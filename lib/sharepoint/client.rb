require "ostruct"
require "ethon"
require "byebug"
require "uri"
require "json"
require "time"

module Sharepoint
  class Client
    # The current active client.
    #
    # @return [Sharepoint::Client]
    # @private
    @@client = nil

    # Lazy-initializes and return the current {@@client}.
    #
    # @return [Sharepoint::Client] The current client.
    def self.client
      @@client ||= new
    end

    # Sets the current {@@client} to +client+.
    #
    # @param  [Sharepoint::Client] client The client to set.
    # @return [Sharepoint::Client] The new client.
    def self.client=(client)
      @@client = client
    end

    # @return [OpenStruct] The current configuration.
    attr_reader :config

    # Initializes a new client with given options.
    #
    # @param [Hash] options The client options.
    # @option options [String] :base_uri (DEFAULT_BASE_URI) The API base uri.
    def initialize(config = {})
      @config = OpenStruct.new(config)
      raise "Username Configuration Error" if @config.username.nil? || @config.username == ""
      raise "Password Configuration Error" if @config.password.nil? || @config.password == ""
      raise "Uri Configuration Error"      if @config.uri.nil? || @config.uri == ""

      @user =         @config.username
      @password =     @config.password
      @base_url =     @config.uri
      @base_api_url = "#{@base_url}/_api/web/"
    end

    # Get all the documents from path
    #
    # @params path [String] the path to request the content
    # @return [Array] of OpenStructs with the info of the files in the path
    def self.documents_for path
      client.send("_documents_for", path)
    end

    # Upload a file
    #
    # @param filename [String] the name of the file uploaded
    # @param content [String or IO] the body of the file
    # @param path [String] the path where to upload the file
    # @return [Fixnum] HTTP response code
    def self.upload filename, content, path
      client.send("_upload", filename, content, path)
    end

    # Update metadata of  a file
    #
    # @param filename [String] the name of the file uploaded
    # @param metadata [Hash] the metadata to change
    # @param path [String] the path where to upload the file
    # @return [Fixnum] HTTP response code
    def self.update_metadata filename, metadata, path
      client.send("_update_metadata", filename, metadata, path)
    end

    private

    def _documents_for path
      ethon = ethon_easy_json_requester
      ethon.url = "#{@base_api_url}GetFolderByServerRelativeUrl('#{URI.escape path}')/Files"
      ethon.perform

      raise "Unable to read ERMS folder, received #{ethon.response_code}" unless (200..299).include? ethon.response_code
      threads = []
      rv = []
      result = JSON.parse( ethon.response_body )
      result['d']['results'].each do |file|
        file_struct = OpenStruct.new(
          title: file['Title'],
          path: file['ServerRelativeUrl'],
          name: file['Name'], url: "#{@base_url}#{file['ServerRelativeUrl']}",
          created_at: DateTime.parse(file['TimeCreated']),
          record_type: nil,
          date_of_issue: nil,
        )

        threads << Thread.new {
          ethon2 = ethon_easy_json_requester
          ethon2.url = "#{@base_api_url}GetFileByServerRelativeUrl('#{path}/#{URI.escape file['Name']}')/ListItemAllFields"
          ethon2.perform
          rs = JSON.parse(ethon2.response_body)['d']
          file_struct.record_type = rs['Record_Type']
          file_struct.date_of_issue = rs['Date_of_issue']

          rv << file_struct
        }
      end
      threads.each { |t| t.join }
      rv
    end

    def _upload filename, content, path
      url = "#{@base_api_url}GetFolderByServerRelativeUrl('#{path}')" +
            "/Files/Add(url='#{filename.gsub("'", "`")}',overwrite=true)"
      url = URI.escape(url)
      easy = ethon_easy_json_requester
      easy.headers = { 'accept' => 'application/json;odata=verbose',
                       'X-RequestDigest' => xrequest_digest }
      easy.http_request(url, :post, { body: content })
      easy.perform
      easy.response_code
    end

    def _update_metadata filename, metadata, path
      url = "#{@base_api_url}GetFileByServerRelativeUrl" +
            "('#{path}/#{filename.gsub("'", "`")}')/ListItemAllFields"
      easy = ethon_easy_json_requester
      easy.url = URI.escape(url)
      easy.perform

      update_metadata_url = JSON.parse(easy.response_body)['d']['__metadata']['uri']

      easy = ethon_easy_json_requester
      easy.headers = { 'accept' =>  'application/json;odata=verbose',
                       'content-type' =>  'application/json;odata=verbose',
                       'X-RequestDigest' =>  xrequest_digest,
                       'X-Http-Method' =>  'PATCH',
                       'If-Match' => "*" }
      easy.http_request(update_metadata_url,
                        :post,
                        { body: prepare_metadata(metadata, path) })
      easy.perform
      easy.response_code
    end

    def ethon_easy_json_requester
      easy          = Ethon::Easy.new(httpauth: :ntlm )
      easy.username = @user
      easy.password = @password
      easy.headers  = { 'accept'=> 'application/json;odata=verbose' }
      easy
    end

    def xrequest_digest
      easy = ethon_easy_json_requester
      easy.http_request("#{@base_url}/_api/contextinfo", :post, { body: '' })
      easy.perform
      JSON.parse(easy.response_body)['d']["GetContextWebInformation"]["FormDigestValue"]
    end

    def prepare_metadata(metadata, path)
      easy = ethon_easy_json_requester
      easy.url = URI.escape("#{@base_api_url}GetFolderByServerRelativeUrl('#{path}')")
      easy.perform
      folder_name = JSON.parse(easy.response_body)['d']['Name']

      metadata.inject("{ '__metadata': { 'type': 'SP.Data.#{folder_name.capitalize}Item' }"){ |result, element|
        key = element[0]
        value = element[1].gsub("'", "`")

        result += ", '#{key}': '#{value}'"
      } + " }"
    end
  end
end
