require "ostruct"
require "ethon"
require "uri"
require "json"
require "time"

require 'active_support/core_ext/string/inflections'

module Sharepoint
  class Client
    FILENAME_INVALID_CHARS = ['~','#', '%', '&' , '*', '{', '}',
                              '\\', ':', '<', '>', '?', '/', '|', '"']

    FILENAME_INVALID_CHARS_REGEXP = /[\/\\~#%&*{}:<>?|"]/

    # @return [OpenStruct] The current configuration.
    attr_reader :config

    # Initializes a new client with given options.
    #
    # @param [Hash] options The client options.
    # @return [Sharepoint::Client] client object
    def initialize(config = {})
      @config = OpenStruct.new(config)
      raise Errors::UsernameConfigurationError.new unless string_not_blank?(@config.username)
      raise Errors::PasswordConfigurationError.new unless string_not_blank?(@config.password)
      raise Errors::UriConfigurationError.new      unless valid_config_uri?

      @user         = @config.username
      @password     = @config.password
      @base_url     = @config.uri
      @base_api_url = "#{@base_url}/_api/"
      @base_api_web_url = "#{@base_api_url}web/"
    end

    # Get all the documents from path
    #
    # @params path [String] the path to request the content
    # @return [Array] of OpenStructs with the info of the files in the path
    def documents_for path
      ethon = ethon_easy_json_requester
      ethon.url = "#{@base_api_web_url}GetFolderByServerRelativeUrl('#{URI.escape path}')/Files"
      ethon.perform

      raise "Unable to read ERMS folder, received #{ethon.response_code}" unless (200..299).include? ethon.response_code
      threads = []
      rv = []
      result = JSON.parse( ethon.response_body )
      result['d']['results'].each do |file|
        file_struct = OpenStruct.new(
          title: file['Title'],
          path: file['ServerRelativeUrl'],
          name: file['Name'],
          url: "#{@base_url}#{file['ServerRelativeUrl']}",
          created_at: DateTime.parse(file['TimeCreated']),
          updated_at: DateTime.parse(file['TimeLastModified']),
          record_type: nil,
          date_of_issue: nil,
        )

        threads << Thread.new {
          ethon2 = ethon_easy_json_requester
          ethon2.url = "#{@base_api_web_url}GetFileByServerRelativeUrl('#{URI.escape path}/#{URI.escape file['Name']}')/ListItemAllFields"
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

    # Search for all documents modified from some datetime on.
    # Uses SharePoint Search API endpoint
    #
    # @param datetime [DateTime] some moment in time
    # @param options [Hash] Supported options are:
    #   - list_id: the GUID of the List you want returned documents to belong to
    #   - web_id: the GUID of the Site you want returned documents to belong to
    # @return [Array] of OpenStructs with all properties of search results
    def search_modified_documents datetime, options={}
      ethon = ethon_easy_json_requester
      query = URI.escape build_search_kql_conditions(datetime, options)
      properties = build_search_properties(options)
      ethon.url = "#{@base_api_url}search/query?querytext=#{query}&#{properties}&clienttype='Custom'"
      ethon.perform
      raise "Request failed, received #{ethon.response_code}" unless (200..299).include? ethon.response_code
      parse_search_response(ethon.response_body)
    end

    # Search in a List for all documents modified from some datetime on.
    # Uses OData on Lists API endpoint
    #
    # @param datetime [DateTime] some moment in time
    # @param list_name [String] The name of the SharePoint List you want to
    #        search into. Please note: a Document Library is a List as well.
    # @return [Array] of OpenStructs with all properties of search results
    def list_modified_documents datetime, list_name
      ethon = ethon_easy_json_requester
      date_condition = "Modified ge datetime'#{datetime.utc.iso8601}'"
      document_condition = "FileSystemObjectType eq 0"
      ethon.url = "#{@base_api_web_url}Lists/GetByTitle('#{URI.escape list_name}')/Items?$expand=Folder,File&$filter=#{URI.escape date_condition}&filter=#{URI.escape document_condition}"
      ethon.perform
      raise "Request failed, received #{ethon.response_code}" unless (200..299).include? ethon.response_code
      parse_list_response(ethon.response_body)
    end

    # Upload a file
    #
    # @param filename [String] the name of the file uploaded
    # @param content [String] the body of the file
    # @param path [String] the path where to upload the file
    # @return [Fixnum] HTTP response code
    def upload filename, content, path
      raise Errors::InvalidSharepointFilename.new unless valid_filename? filename

      url = "#{@base_api_web_url}GetFolderByServerRelativeUrl('#{path}')" +
            "/Files/Add(url='#{filename.gsub("'", "`")}',overwrite=true)"
      url = URI.escape(url)
      easy = ethon_easy_json_requester
      easy.headers = { 'accept' => 'application/json;odata=verbose',
                       'X-RequestDigest' => xrequest_digest }
      easy.http_request(url, :post, { body: content })
      easy.perform
      easy.response_code
    end

    # Update metadata of  a file
    #
    # @param filename [String] the name of the file uploaded
    # @param metadata [Hash] the metadata to change
    # @param path [String] the path where to upload the file
    # @return [Fixnum] HTTP response code
    def update_metadata filename, metadata, path
      prepared_metadata = prepare_metadata(metadata, path)

      url = "#{@base_api_web_url}GetFileByServerRelativeUrl" +
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
                        { body: prepared_metadata })
      easy.perform
      easy.response_code
    end

    private

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
      easy.url = URI.escape("#{@base_api_web_url}GetFolderByServerRelativeUrl('#{path}')")
      easy.perform
      folder_name = JSON.parse(easy.response_body)['d']['Name']

      metadata.inject("{ '__metadata': { 'type': 'SP.Data.#{folder_name.capitalize}Item' }"){ |result, element|
        key = element[0]
        value = element[1]

        raise Errors::InvalidMetadata.new if key.to_s.include?("'") || value.include?("'")

        result += ", '#{key}': '#{value}'"
      } + " }"
    end

    def string_not_blank?(object)
      !object.nil? && object != "" && object.is_a?(String)
    end

    def valid_config_uri?
      if @config.uri and @config.uri.is_a? String
        uri = URI.parse(@config.uri)
        uri.kind_of?(URI::HTTP) || uri.kind_of?(URI::HTTPS)
      else
        false
      end
    end

    def valid_filename? name
      (name =~ FILENAME_INVALID_CHARS_REGEXP).nil?
    end

    def build_search_kql_conditions(datetime, options)
      conditions = []
      conditions << "write>=#{datetime.utc.iso8601}"
      conditions << "IsDocument=1"
      conditions << "WebId=#{options[:web_id]}" unless options[:web_id].nil?
      conditions << "ListId:#{options[:list_id]}" unless options[:list_id].nil?
      "'#{conditions.join('+')}'"
    end

    def build_search_properties(options)
      default_properties = %w(
        Write IsDocument ListId WebId
        Created Title Author Size Path
      )
      properties = options[:properties] || []
      properties += default_properties
      "selectproperties='#{properties.join(',')}'"
    end

    def parse_search_response(response_body)
      json_response = JSON.parse(response_body)
      search_results = json_response.dig('d', 'query', 'PrimaryQueryResult', 'RelevantResults', 'Table', 'Rows', 'results')
      records = []
      search_results.each do |result|
        record = {}
        result.dig('Cells', 'results').each do |result_attrs|
          key = result_attrs['Key'].underscore.to_sym
          record[key] = result_attrs['Value']
        end
        records << OpenStruct.new(record)
      end
      records
    end

    def parse_list_response(response_body)
      json_response = JSON.parse(response_body)
      results = json_response['d']['results']
      records = []
      results.each do |result|
        record = {}
        %w( GUID Created Modified ).each do |key|
          record[key.underscore.to_sym] = result[key]
        end
        file = result['File']
        %w( Name ServerRelativeUrl Title ).each do |key|
          record[key.underscore.to_sym] = file[key]
        end
        records << OpenStruct.new(record)
      end
      records
    end

  end
end
