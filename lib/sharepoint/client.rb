# frozen_string_literal: true

require 'ostruct'
require 'ethon'
require 'uri'
require 'json'
require 'time'

require 'active_support/core_ext/string/inflections'
require 'active_support/core_ext/object/blank'

require_relative 'client/token'

module Sharepoint
  class Client
    FILENAME_INVALID_CHARS = '~"#%&*:<>?/\{|}'

    DEFAULT_TOKEN_ETHON_OPTIONS = { followlocation: 1, maxredirs: 5 }.freeze
    VALID_TOKEN_CONFIG_OPTIONS = %i[client_id client_secret tenant_id cert_name auth_scope].freeze

    DEFAULT_NTLM_ETHON_OPTIONS = { httpauth: :ntlm, followlocation: 1, maxredirs: 5 }.freeze
    VALID_NTLM_CONFIG_OPTIONS = %i[username password].freeze

    URI_PARSER =
      if defined?(URI::RFC2396_PARSER)
        URI::RFC2396_PARSER
      else
        URI::DEFAULT_PARSER
      end
    private_constant :URI_PARSER

    def authenticating_with_token
      generate_new_token
      yield
    end

    def generate_new_token
      token.retrieve
    end

    def bearer_auth
      "Bearer #{token}"
    end

    # @return [OpenStruct] The current configuration.
    attr_reader :config
    attr_reader :token

    # Initializes a new client with given options.
    #
    # @param [Hash] config The client options:
    #  - `:uri` The SharePoint server's root url
    #  - `:authentication` The authentication method to use [:ntlm, :token]
    #  - `:username` self-explanatory
    #  - `:password` self-explanatory
    #  - `:client_id` self-explanatory
    #  - `:client_secret` self-explanatory
    #  - `:tenant_id` self-explanatory
    #  - `:cert_name` self-explanatory
    #  - `:auth_scope` self-explanatory
    # @return [Sharepoint::Client] client object
    def initialize(config = {})
      @config = OpenStruct.new(config)
      @token = Token.new(@config)
      validate_config!
    end

    # Get all the documents from path
    #
    # @param path [String] the path to request the content
    # @param site_path [String] if the SP instance contains sites, the site path, e.g. "/sites/my-site"
    #
    # @return [Array] of OpenStructs with the info of the files in the path
    def documents_for(path, site_path = '')
      ethon = ethon_easy_json_requester
      ethon.url = "#{computed_web_api_url(site_path)}GetFolderByServerRelativeUrl('#{uri_escape path}')/Files"
      ethon.perform
      check_and_raise_failure(ethon)

      threads = []
      rv = []
      result = JSON.parse(ethon.response_body)
      result['d']['results'].each do |file|
        file_struct = OpenStruct.new(
          title: file['Title'],
          path: file['ServerRelativeUrl'],
          name: file['Name'],
          url: "#{base_url}#{file['ServerRelativeUrl']}",
          created_at: Time.parse(file['TimeCreated']),
          updated_at: Time.parse(file['TimeLastModified']),
          record_type: nil,
          date_of_issue: nil
        )

        threads << Thread.new do
          ethon2 = ethon_easy_json_requester
          server_relative_url = "#{site_path}#{path}/#{file['Name']}"
          ethon2.url = "#{computed_web_api_url(site_path)}GetFileByServerRelativeUrl('#{uri_escape server_relative_url}')/ListItemAllFields"
          ethon2.perform
          rs = JSON.parse(ethon2.response_body)['d']
          file_struct.record_type = rs['Record_Type']
          file_struct.date_of_issue = rs['Date_of_issue']

          rv << file_struct
        end
      end
      threads.each { |t| t.join }
      rv
    end

    # Checks whether a document exists with the given path
    #
    # @param file_path [String] the file path, without the site path if any
    # @param site_path [String] if the SP instance contains sites, the site path, e.g. "/sites/my-site"
    #
    # @return `true` if document exists, false otherwise.
    def document_exists?(file_path, site_path = nil)
      file = split_path(file_path)
      sanitized_filename = sanitize_filename(file[:name])
      server_relative_url = "#{site_path}#{file[:path]}/#{sanitized_filename}"
      url = computed_web_api_url(site_path)
      ethon = ethon_easy_json_requester
      ethon.url = uri_escape "#{url}GetFileByServerRelativeUrl('#{odata_escape_single_quote server_relative_url}')"
      ethon.perform
      exists = false
      if ethon.response_code.eql? 200
        json_response = JSON.parse(ethon.response_body)
        if json_response['d'] &&
           json_response['d']['ServerRelativeUrl'].eql?(server_relative_url)
          exists = true
        end
      end
      exists
    end

    # Get a document's metadata
    #
    # @param file_path [String] the file path, without the site path if any
    # @param site_path [String] if the SP instance contains sites, the site path, e.g. "/sites/my-site"
    # @param custom_properties [Array] of String with names of custom properties to be returned
    #
    # @return [OpenStruct] with both default and custom metadata
    def get_document(file_path, site_path = nil, custom_properties = [])
      url = computed_web_api_url(site_path)
      server_relative_url = odata_escape_single_quote "#{site_path}#{file_path}"
      ethon = ethon_easy_json_requester
      ethon.url = "#{url}GetFileByServerRelativeUrl('#{uri_escape server_relative_url}')/ListItemAllFields"
      ethon.perform
      check_and_raise_failure(ethon)
      parse_get_document_response(ethon.response_body, custom_properties)
    end

    # Search for all documents modified in a given time range,
    # boundaries included. Uses SharePoint Search API endpoint
    #
    # @param options [Hash] Supported options are:
    #   * start_at [Time] Range start time (mandatory)
    #   * end_at [Time] Range end time (optional). If null, documents modified
    #     after start_at will be returned
    #   * list_id [String] the GUID of the List you want returned documents
    #     to belong to (optional)
    #   * web_id [String] the GUID of the Site you want returned documents
    #     to belong to (optional)
    #   * properties [Array] of String with names of custom properties
    #     to be returned (optional)
    #   * max_results [Integer] the maximum number of results to be returned;
    #     defaults to 500 which is the default `MaxRowLimit` in SharePoint 2013.
    #     If you have increased that in your on-premise SP instance, then that's
    #     your limit for `max_results` param as well
    #   * start_result [Integer] the offset for results to be returned; defaults to 0.
    #     Useful when more than `max_results` documents have been modified in your
    #     time range, so you can iterate to fetch'em all.
    #
    # @return [Hash] with the following keys:
    #   * `:requested_url` [String] the URL requested to the SharePoint server
    #   * `:server_responded_at` [Time] the time when server returned its response
    #   * `:results` [Array] of OpenStructs with all properties of search results,
    #      sorted by last modified date (`write`)
    def search_modified_documents(options = {})
      ethon = ethon_easy_json_requester
      query = uri_escape build_search_kql_conditions(options)
      properties = build_search_properties(options)
      filters = build_search_fql_conditions(options)
      sorting = "sortlist='write:ascending'"
      paging = build_search_paging(options)
      ethon.url = "#{base_api_url}search/query?querytext=#{query}&refinementfilters=#{filters}&#{properties}&#{sorting}&#{paging}&clienttype='Custom'"
      ethon.perform
      check_and_raise_failure(ethon)
      server_responded_at = Time.now
      {
        requested_url: ethon.url,
        server_responded_at: server_responded_at,
        results: parse_search_response(ethon.response_body)
      }
    end

    # Dumb wrapper of SharePoint Search API endpoint.
    #
    # @param options [Hash] All key => values in this hash will be passed to
    #  the `/search/query` endpoint as param=value in the querystring.
    #  Some very useful ones are:
    #  * `:querytext` [String] A valid KQL query. See:
    #    https://msdn.microsoft.com/en-us/library/office/ee558911.aspx
    #  * `:refinementfilters` [String] A valid query using OData syntax. See:
    #    https://msdn.microsoft.com/en-us/library/office/fp142385.aspx
    #  * `:selectProperties` [String] A comma-separated list of properties
    #    whose values you want returned for your results
    #  * `:rowlimit` [Number] The number of results to be returned (max 500)
    # @return [Hash] with the following keys:
    #   * `:requested_url` [String] the URL requested to the SharePoint server
    #   * `:server_responded_at` [Time] the time when server returned its response
    #   * `:results` [Array] of OpenStructs with all properties of search results
    def search(options = {})
      params = []
      options.each do |key, value|
        params << "#{key}=#{value}"
      end
      ethon = ethon_easy_json_requester
      ethon.url = uri_escape("#{base_api_url}search/query?#{params.join('&')}")
      ethon.perform
      check_and_raise_failure(ethon)
      server_responded_at = Time.now
      {
        requested_url: ethon.url,
        server_responded_at: server_responded_at,
        results: parse_search_response(ethon.response_body)
      }
    end

    # Search in a List for all documents matching the passed conditions.
    #
    # @param list_name [String] The name of the SharePoint List you want to
    #        search into. Please note: a Document Library is a List as well.
    # @param conditions [String] OData conditions that returned documents
    #        should verify, or nil if you want all documents. See:
    #        https://msdn.microsoft.com/en-us/library/office/fp142385.aspx
    # @param site_path [String] if the SP instance contains sites, the site path,
    #        e.g. "/sites/my-site"
    # @param properties [Array] of String with names of custom properties to be returned
    #
    # @return [Hash] with the following keys:
    #   * `:requested_url` [String] the URL requested to the SharePoint server
    #   * `:server_responded_at` [Time] the time when server returned its response
    #   * `:results` [Array] of OpenStructs with all properties of search results
    def list_documents(list_name, conditions, site_path = nil, properties = [])
      url = computed_web_api_url(site_path)
      filter_param = "$filter=#{conditions}" if conditions.present?
      expand_param = '$expand=Folder,File'
      default_properties = %w[FileSystemObjectType UniqueId Title Created Modified File]
      all_properties = default_properties + properties
      select_param = "$select=#{all_properties.join(',')}"
      url = "#{url}Lists/GetByTitle('#{odata_escape_single_quote(list_name)}')/Items?#{expand_param}&#{select_param}"
      url += "&#{filter_param}"

      records = []
      page_url = uri_escape url
      loop do
        body = list_documents_page(page_url)
        records += parse_list_response(body, all_properties)
        page_url = body['d']['__next']
        break if page_url.blank?
      end

      server_responded_at = Time.now

      {
        requested_url: url,
        server_responded_at: server_responded_at,
        results: records
      }
    end

    def list_documents_page(url)
      ethon = ethon_easy_json_requester
      ethon.url = url
      ethon.perform
      check_and_raise_failure(ethon)

      JSON.parse(ethon.response_body)
    end

    # Get a document's file contents. If it's a link to another document, it's followed.
    #
    # @param file_path [String] the file path, without the site path if any
    # @param site_path [String] if the SP instance contains sites, the site path, e.g. "/sites/my-site"
    # @param link_credentials [Hash] credentials to access the link's destination repo.
    # Accepted keys: `:username` and `:password`
    #
    # @return [Hash] with the following keys:
    #  - `:file_contents` [String] with the file contents
    #  - `:link_url` [String] if the requested file is a link, this returns the destination file url
    def download(file_path: nil, site_path: nil, link_credentials: {})
      meta = get_document(file_path, site_path)
      meta_path = meta.url || meta.path
      if meta_path.nil?
        url = computed_web_api_url(site_path)
        server_relative_url = odata_escape_single_quote "#{site_path}#{file_path}"
        download_url "#{url}GetFileByServerRelativeUrl('#{server_relative_url}')/$value"
      else # requested file is a link
        paths = extract_paths(meta_path)
        link_config = { uri: paths[:root] }
        if link_credentials.empty?
          link_config = config.to_h.merge(link_config)
        else
          link_config.merge!(link_credentials)
        end
        link_client = self.class.new(link_config)
        result = link_client.download_url meta_path
        result[:link_url] = meta.url if result[:link_url].nil?
        result
      end
    end

    # Downloads a file provided its full URL. Follows redirects.
    #
    # @param url [String] the URL of the file to download

    # @return [Hash] with the following keys:
    #  - `:file_contents` [String] with the file contents
    #  - `:link_url` [String] if some redirect is followed, returns the last `Location:` header value
    def download_url(url)
      ethon = ethon_easy_requester
      ethon.url = uri_escape(url)
      ethon.perform
      check_and_raise_failure(ethon)
      {
        file_contents: ethon.response_body,
        link_url: last_location_header(ethon)
      }
    end

    # Creates a folder
    #
    # @param name [String] the name of the folder
    # @param path [String] the path where to create the folder
    # @param site_path [String] if the SP instance contains sites, the site path, e.g. "/sites/my-site"
    #
    # @return [Fixnum] HTTP response code
    def create_folder(name, path, site_path = nil)
      return unless name

      sanitized_name = sanitize_filename(name)
      url = computed_web_api_url(site_path)
      path = path[1..] if path[0].eql?('/')
      url = uri_escape "#{url}GetFolderByServerRelativeUrl('#{path}')/Folders"
      easy = ethon_easy_json_requester
      easy.headers = with_bearer_authentication_header({
                                                         'accept' => 'application/json;odata=verbose',
                                                         'content-type' => 'application/json;odata=verbose',
                                                         'X-RequestDigest' => xrequest_digest(site_path)
                                                       })
      payload = {
        '__metadata' => {
          'type' => 'SP.Folder'
        },
        'ServerRelativeUrl' => "#{path}/#{sanitized_name}"
      }
      easy.http_request(url, :post, body: payload.to_json)
      easy.perform
      check_and_raise_failure(easy)
      easy.response_code
    end

    # Checks if a folder exists
    #
    # @param path [String] the folder path
    # @param site_path [String] if the SP instance contains sites, the site path, e.g. "/sites/my-site"
    #
    # @return [Fixnum] HTTP response code
    def folder_exists?(path, site_path = nil)
      url = computed_web_api_url(site_path)
      path = [site_path, path].compact.join('/')
      url = uri_escape "#{url}GetFolderByServerRelativeUrl('#{path}')"
      easy = ethon_easy_json_requester
      easy.http_request(url, :get)
      easy.perform
      easy.response_code == 200
    end

    # Upload a file
    #
    # @param filename [String] the name of the file uploaded
    # @param content [String] the body of the file
    # @param path [String] the path where to upload the file
    # @param site_path [String] if the SP instance contains sites, the site path, e.g. "/sites/my-site"
    #
    # @return [Fixnum] HTTP response code
    def upload(filename, content, path, site_path = nil)
      sanitized_filename = sanitize_filename(filename)
      url = computed_web_api_url(site_path)
      path = path[1..] if path[0].eql?('/')
      url = uri_escape "#{url}GetFolderByServerRelativeUrl('#{path}')/Files/Add(url='#{sanitized_filename}',overwrite=true)"
      easy = ethon_easy_json_requester
      easy.headers = with_bearer_authentication_header({ 'accept' => 'application/json;odata=verbose',
                                                         'X-RequestDigest' => xrequest_digest(site_path) })
      easy.http_request(url, :post, { body: content })
      easy.perform
      check_and_raise_failure(easy)
      easy.response_code
    end

    # Update metadata of  a file
    #
    # @param filename [String] the name of the file
    # @param metadata [Hash] the metadata to change
    # @param path [String] the path where the file is stored
    # @param site_path [String] if the SP instance contains sites, the site path, e.g. "/sites/my-site"
    #
    # @return [Fixnum] HTTP response code
    def update_metadata(filename, metadata, path, site_path = nil)
      sanitized_filename = sanitize_filename(filename)
      url = computed_web_api_url(site_path)
      server_relative_url = "#{site_path}#{path}/#{sanitized_filename}"
      easy = ethon_easy_json_requester
      easy.url = uri_escape "#{url}GetFileByServerRelativeUrl('#{server_relative_url}')/ListItemAllFields"
      easy.perform

      __metadata = JSON.parse(easy.response_body)['d']['__metadata']
      update_metadata_url = __metadata['uri']
      prepared_metadata = prepare_metadata(metadata, __metadata['type'])

      easy = ethon_easy_json_requester
      easy.headers = with_bearer_authentication_header({ 'accept' => 'application/json;odata=verbose',
                                                         'content-type' => 'application/json;odata=verbose',
                                                         'X-RequestDigest' => xrequest_digest(site_path),
                                                         'X-Http-Method' => 'PATCH',
                                                         'If-Match' => '*' })
      easy.http_request(update_metadata_url,
                        :post,
                        { body: prepared_metadata })
      easy.perform
      check_and_raise_failure(easy)
      easy.response_code
    end

    # Search for all lists in the SP instance
    #
    # @param site_path [String] if the SP instance contains sites, the site path, e.g. "/sites/my-site"
    # @param query [Hash] Hash with OData query operations, e.g. `{ select: 'Id,Title', filter: 'ItemCount gt 0 and Hidden eq false' }`.
    #
    # @return [Hash] with the following keys:
    #   * `:requested_url` [String] the URL requested to the SharePoint server
    #   * `:server_responded_at` [Time] the time when server returned its response
    #   * `:results` [Array] of OpenStructs with all lists returned by the query
    def lists(site_path = '', query = {})
      url = "#{computed_web_api_url(site_path)}Lists"
      url << "?#{build_query_params(query)}" if query.present?

      ethon = ethon_easy_json_requester
      ethon.url = uri_escape(url)
      ethon.perform
      check_and_raise_failure(ethon)

      {
        requested_url: ethon.url,
        server_responded_at: Time.now,
        results: parse_lists_in_site_response(ethon.response_body)
      }
    end

    # Returns a list of resource
    #
    # @param list_name [String] the name of the list
    # @param fields [Array][String] fields to narrow down the list content
    # @param site_path [String] if the SP instance contains sites, the site path, e.g. "/sites/my-site"
    #
    # @return [Fixnum] HTTP response code
    def index(list_name, site_path = '', fields = [])
      url = computed_web_api_url(site_path)
      url = "#{url}lists/GetByTitle('#{odata_escape_single_quote(list_name)}')/items"
      url += "?$select=#{fields.join(',')}" if fields

      process_url(uri_escape(url), fields)
    end

    # Index a list field. Requires admin permissions
    #
    # @param list_name [String] the name of the list
    # @param field_name [String] the name of the field to index
    # @param site_path [String] if the SP instance contains sites, the site path, e.g. "/sites/my-site"
    #
    # @return [Fixnum] HTTP response code
    def index_field(list_name, field_name, site_path = '')
      url = computed_web_api_url(site_path)
      easy = ethon_easy_json_requester
      easy.url = uri_escape "#{url}Lists/GetByTitle('#{odata_escape_single_quote(list_name)}')/Fields/getByTitle('#{field_name}')"
      easy.perform

      parsed_response_body = JSON.parse(easy.response_body)
      return 304 if parsed_response_body['d']['Indexed']

      update_object_metadata parsed_response_body['d']['__metadata'], { 'Indexed' => true }, site_path
    end

    def requester
      ethon_easy_requester
    end

    # ========================================================================
    # COLLABORATION METHODS
    # ========================================================================
    # The following methods support SharePoint collaboration workflows for
    # document sharing, permission management, and web editing.
    #
    # Configuration requirements for collaboration features:
    #   - site_path: The SharePoint site path (e.g., "/sites/my-site")
    #   - base_folder: Base folder for collaboration documents
    #   - base_uri: SharePoint base URI (e.g., "https://company.sharepoint.com")
    # ========================================================================

    FOLDER_SEPARATOR = '/'

    OFFICE_TYPE_EXTENSIONS = {
      '.doc' => 'w',   # Word
      '.docx' => 'w',  # Word
      '.xls' => 'x',   # Excel
      '.xlsx' => 'x',  # Excel
      '.ppt' => 'p',   # PowerPoint
      '.pptx' => 'p'   # PowerPoint
    }.freeze

    # Uploads a document to SharePoint and returns its UniqueId
    #
    # @param filename [String] The name of the file
    # @param file_content [String, IO] The file content
    # @param folder_path [String] The relative folder path within base_folder
    # @param metadata [Hash] Optional metadata for the document
    # @return [Hash] { success: true, unique_id: String } or raises error
    # @raise [UploadError] if upload fails
    def collaboration_upload_document(filename, file_content, folder_path, metadata = {})
      full_path = collaboration_build_full_path(folder_path)

      collaboration_ensure_folder_exists(folder_path)

      # Upload and capture the response with UniqueId
      unique_id = collaboration_upload_and_get_unique_id(filename, file_content, full_path)

      update_metadata(filename, metadata, full_path, collaboration_site_path) unless metadata.empty?

      { success: true, unique_id: unique_id }
    rescue StandardError => e
      raise Sharepoint::Errors::UploadError.new("Failed to upload document: #{e.message}")
    end

    # Grants permissions to multiple users for a document
    #
    # @param filename [String] The name of the file
    # @param folder_path [String] The relative folder path within base_folder
    # @param user_emails [Array<String>] Array of user email addresses
    # @param role_id [Integer] SharePoint role definition ID (e.g., 1073741827 for Contribute)
    # @return [Boolean] true if successful
    # @raise [PermissionError] if permission grant fails
    def collaboration_grant_permissions(filename, folder_path, user_emails, role_id)
      raise Sharepoint::Errors::PermissionError.new('user_emails cannot be nil') if user_emails.nil?
      raise Sharepoint::Errors::PermissionError.new('user_emails must be an array') unless user_emails.is_a?(Array)
      raise Sharepoint::Errors::PermissionError.new('user_emails cannot be empty') if user_emails.empty?
      raise Sharepoint::Errors::PermissionError.new('role_id is required') if role_id.nil?

      server_relative_url = collaboration_build_server_relative_url(filename, folder_path)

      # First, ensure users are added to the site and break role inheritance
      user_ids = user_emails.compact.map { |email| collaboration_ensure_user_and_get_id(email) }
      collaboration_break_role_inheritance(server_relative_url)

      # Grant permissions to each user
      user_ids.each do |user_id|
        collaboration_grant_user_permission(server_relative_url, user_id, role_id)
      end

      true
    rescue Sharepoint::Errors::PermissionError
      raise
    rescue StandardError => e
      raise Sharepoint::Errors::PermissionError.new("Failed to grant permissions: #{e.message}\n#{e.backtrace.first(5).join("\n")}")
    end

    # Revokes a single user's permission from a document
    #
    # @param filename [String] The name of the file
    # @param folder_path [String] The relative folder path within base_folder
    # @param user_email [String] The user's email address
    # @return [Boolean] true if successful
    # @raise [PermissionError] if permission revocation fails
    def collaboration_revoke_user_permission(filename, folder_path, user_email)
      server_relative_url = collaboration_build_server_relative_url(filename, folder_path)
      user_id = collaboration_get_user_id(user_email)

      collaboration_revoke_permission(server_relative_url, user_id)

      true
    rescue StandardError => e
      raise Sharepoint::Errors::PermissionError.new("Failed to revoke user permission: #{e.message}")
    end

    # Revokes all permissions from a document (resets to inherited permissions)
    #
    # @param filename [String] The name of the file
    # @param folder_path [String] The relative folder path within base_folder
    # @return [Boolean] true if successful
    # @raise [PermissionError] if permission reset fails
    def collaboration_revoke_all_permissions(filename, folder_path)
      server_relative_url = collaboration_build_server_relative_url(filename, folder_path)

      collaboration_reset_role_inheritance(server_relative_url)

      true
    rescue StandardError => e
      raise Sharepoint::Errors::PermissionError.new("Failed to revoke all permissions: #{e.message}")
    end

    # Generates the web edit URL for a document
    #
    # @param filename [String] The name of the file
    # @param folder_path [String] The relative folder path within base_folder
    # @param unique_id [String] Optional UniqueId (GUID) of the file
    # @return [String] The web edit URL
    def collaboration_get_web_edit_url(filename, folder_path, unique_id = nil)
      # Use provided unique_id if available, otherwise try to fetch it
      file_unique_id = unique_id || begin
        server_relative_url = collaboration_build_server_relative_url(filename, folder_path)
        collaboration_get_file_unique_id(server_relative_url)
      end

      office_type = detect_office_type(filename)

      if file_unique_id.present?
        # Use Doc.aspx with GUID format for web editing (most reliable)
        escaped_filename = CGI.escape(filename)
        "#{collaboration_base_uri}/:#{office_type}:/r#{collaboration_site_path}/_layouts/15/Doc.aspx?sourcedoc=%7B#{file_unique_id.upcase}%7D&file=#{escaped_filename}&action=default&mobileredirect=true&DefaultItemOpen=1"
      else
        # Fallback: Link to the folder view where user can click the file
        escaped_folder = CGI.escape(folder_path)
        "#{collaboration_base_uri}#{collaboration_site_path}/#{collaboration_base_folder}/#{escaped_folder}"
      end
    end

    # Detects the Office file type for SharePoint web URLs
    #
    # @param filename [String] The filename
    # @return [String] The office type code (w, x, p, or d)
    def detect_office_type(filename)
      extension = File.extname(filename).downcase
      OFFICE_TYPE_EXTENSIONS.fetch(extension, 'd')
    end

    # Gets the UniqueId (GUID) of a file
    #
    # @param server_relative_url [String] The server-relative URL of the file
    # @return [String, nil] The file's UniqueId or nil if not found
    def collaboration_get_file_unique_id(server_relative_url)
      ethon = ethon_easy_requester
      escaped_url = CGI.escape(server_relative_url)
      ethon.url = "#{collaboration_base_uri}/_api/web/getfilebyserverrelativeurl('#{escaped_url}')?$select=UniqueId"
      ethon.headers = { 'Accept' => 'application/json;odata=verbose' }

      authenticating_with_token do
        ethon.perform
      end

      if defined?(Rails) && Rails.logger
        Rails.logger.info("SharePoint UniqueId API response code: #{ethon.response_code}")
        Rails.logger.info("SharePoint UniqueId API URL: #{ethon.url}")
      end

      if ethon.response_code == 200
        response = JSON.parse(ethon.response_body)
        unique_id = response.dig('d', 'UniqueId')
        Rails.logger.info("SharePoint UniqueId retrieved: #{unique_id}") if defined?(Rails) && Rails.logger
        unique_id
      else
        Rails.logger.warn("SharePoint UniqueId API failed with code #{ethon.response_code}: #{ethon.response_body[0..500]}") if defined?(Rails) && Rails.logger
        nil
      end
    rescue StandardError => e
      if defined?(Rails) && Rails.logger
        Rails.logger.error("Failed to get file UniqueId: #{e.message}")
        Rails.logger.error(e.backtrace[0..5].join("\n"))
      end
      nil
    end

    # Downloads a document from SharePoint
    #
    # @param filename [String] The name of the file
    # @param folder_path [String] The relative folder path within base_folder
    # @return [String] The file content
    # @raise [DownloadError] if download fails
    def collaboration_download_document(filename, folder_path)
      full_path = collaboration_build_full_path(folder_path)
      # The download method expects the file_path without site_path, with leading slash
      file_path = "/#{full_path}/#{filename}"

      # Use the existing download method which handles authentication properly
      result = download(file_path: file_path, site_path: collaboration_site_path)

      # The download method returns a hash with :file_contents key
      result.is_a?(Hash) ? result[:file_contents] : result
    rescue StandardError => e
      raise Sharepoint::Errors::DownloadError.new("Failed to download document: #{e.message}")
    end

    # Deletes a document from SharePoint
    #
    # @param filename [String] The name of the file
    # @param folder_path [String] The relative folder path within base_folder
    # @return [Boolean] true if successful
    # @raise [DeleteError] if deletion fails
    def collaboration_delete_document(filename, folder_path)
      server_relative_url = collaboration_build_server_relative_url(filename, folder_path)

      ethon = ethon_easy_json_requester
      ethon.headers = with_bearer_authentication_header({
                                                          'accept' => 'application/json;odata=verbose',
                                                          'X-RequestDigest' => xrequest_digest(collaboration_site_path),
                                                          'IF-MATCH' => '*',
                                                          'X-HTTP-Method' => 'DELETE'
                                                        })

      authenticating_with_token do
        ethon.http_request(
          uri_escape("#{collaboration_base_uri}#{collaboration_site_path}/_api/web/GetFileByServerRelativeUrl('#{server_relative_url}')"),
          :post,
          body: ''
        )
        ethon.perform
      end

      raise Sharepoint::Errors::DeleteError.new("Failed to delete document: HTTP #{ethon.response_code}") unless ethon.response_code == 200

      true
    rescue Sharepoint::Errors::DeleteError
      raise
    rescue StandardError => e
      raise Sharepoint::Errors::DeleteError.new("Failed to delete document: #{e.message}")
    end

    # Checks if a document exists in SharePoint (collaboration version)
    #
    # @param filename [String] The name of the file
    # @param folder_path [String] The relative folder path within base_folder
    # @return [Boolean] true if document exists
    def collaboration_document_exists?(filename, folder_path)
      server_relative_url = collaboration_build_server_relative_url(filename, folder_path)

      ethon = ethon_easy_json_requester
      ethon.url = uri_escape("#{collaboration_base_uri}#{collaboration_site_path}/_api/web/GetFileByServerRelativeUrl('#{server_relative_url}')")

      authenticating_with_token do
        ethon.perform
      end

      ethon.response_code == 200
    rescue StandardError
      false
    end

    # Deletes a folder and all its contents from SharePoint
    #
    # @param folder_path [String] The relative folder path within base_folder
    # @return [Boolean] true if successful
    # @raise [DeleteError] if deletion fails
    def collaboration_delete_folder(folder_path)
      full_path = collaboration_build_full_path(folder_path)
      server_relative_url = "#{collaboration_site_path}/#{full_path}"

      ethon = ethon_easy_json_requester
      ethon.headers = with_bearer_authentication_header({
                                                          'accept' => 'application/json;odata=verbose',
                                                          'X-RequestDigest' => xrequest_digest(collaboration_site_path),
                                                          'IF-MATCH' => '*',
                                                          'X-HTTP-Method' => 'DELETE'
                                                        })

      authenticating_with_token do
        ethon.http_request(
          uri_escape("#{collaboration_base_uri}#{collaboration_site_path}/_api/web/GetFolderByServerRelativeUrl('#{server_relative_url}')"),
          :post,
          body: ''
        )
        ethon.perform
      end

      unless ethon.response_code == 200
        raise Sharepoint::Errors::DeleteError.new(
          "Failed to delete folder: HTTP #{ethon.response_code} - #{ethon.response_body}"
        )
      end

      true
    rescue Sharepoint::Errors::DeleteError
      raise
    rescue StandardError => e
      raise Sharepoint::Errors::DeleteError.new("Failed to delete folder: #{e.message}")
    end

    private

    # Gets collaboration configuration values with fallback
    def collaboration_site_path
      @config.site_path || ''
    end

    def collaboration_base_folder
      @config.base_folder || ''
    end

    def collaboration_base_uri
      @config.base_uri || @config.uri
    end

    # Builds the full path within SharePoint
    def collaboration_build_full_path(folder_path)
      [collaboration_base_folder, folder_path].reject(&:empty?).join(FOLDER_SEPARATOR)
    end

    # Builds the server-relative URL for a document
    def collaboration_build_server_relative_url(filename, folder_path)
      full_path = collaboration_build_full_path(folder_path)
      "#{collaboration_site_path}/#{full_path}/#{filename}"
    end

    # Ensures a folder exists, creating it if necessary
    def collaboration_ensure_folder_exists(folder_path)
      full_path = collaboration_build_full_path(folder_path)
      return if folder_exists?(full_path, collaboration_site_path)

      create_folder(folder_path, collaboration_base_folder, collaboration_site_path)
    end

    # Uploads a file and extracts UniqueId from response
    def collaboration_upload_and_get_unique_id(filename, file_content, full_path)
      sanitized_filename = sanitize_filename(filename)
      url = computed_web_api_url(collaboration_site_path)
      path = full_path[0].eql?('/') ? full_path[1..] : full_path
      upload_url = uri_escape("#{url}GetFolderByServerRelativeUrl('#{path}')/Files/Add(url='#{sanitized_filename}',overwrite=true)")

      ethon = ethon_easy_json_requester
      ethon.headers = with_bearer_authentication_header({
                                                          'accept' => 'application/json;odata=verbose',
                                                          'X-RequestDigest' => xrequest_digest(collaboration_site_path)
                                                        })
      ethon.http_request(upload_url, :post, { body: file_content })

      authenticating_with_token do
        ethon.perform
      end

      check_and_raise_failure(ethon)

      if ethon.response_code == 200
        begin
          response = JSON.parse(ethon.response_body)
          unique_id = response.dig('d', 'UniqueId')
          Rails.logger.info("SharePoint upload successful, UniqueId: #{unique_id}") if defined?(Rails) && Rails.logger
          unique_id
        rescue JSON::ParserError => e
          Rails.logger.error("Failed to parse upload response: #{e.message}") if defined?(Rails) && Rails.logger
          nil
        end
      else
        Rails.logger.warn("SharePoint upload returned code #{ethon.response_code}") if defined?(Rails) && Rails.logger
        nil
      end
    rescue StandardError => e
      Rails.logger.error("Failed to get UniqueId from upload: #{e.message}") if defined?(Rails) && Rails.logger
      nil
    end

    # Adds a user to the site and returns their ID
    def collaboration_ensure_user_and_get_id(email)
      payload = { 'logonName' => "i:0#.f|membership|#{email}" }

      ethon = ethon_easy_json_requester
      ethon.headers = with_bearer_authentication_header({
                                                          'accept' => 'application/json;odata=verbose',
                                                          'content-type' => 'application/json;odata=verbose',
                                                          'X-RequestDigest' => xrequest_digest(collaboration_site_path)
                                                        })

      authenticating_with_token do
        ethon.http_request(
          uri_escape("#{collaboration_base_uri}#{collaboration_site_path}/_api/web/ensureuser"),
          :post,
          body: payload.to_json
        )
        ethon.perform
      end

      raise Sharepoint::Errors::PermissionError.new("Failed to ensure user #{email}: HTTP #{ethon.response_code}, Response: #{ethon.response_body}") unless ethon.response_code == 200

      response = JSON.parse(ethon.response_body)
      user_id = response.dig('d', 'Id')

      raise Sharepoint::Errors::PermissionError.new("Failed to get user ID for #{email}") if user_id.nil?

      user_id
    end

    # Gets a user ID from their email (assumes they're already in the site)
    def collaboration_get_user_id(email)
      collaboration_ensure_user_and_get_id(email)
    end

    # Breaks role inheritance for a document
    def collaboration_break_role_inheritance(server_relative_url)
      ethon = ethon_easy_json_requester
      ethon.headers = with_bearer_authentication_header({
                                                          'accept' => 'application/json;odata=verbose',
                                                          'X-RequestDigest' => xrequest_digest(collaboration_site_path)
                                                        })

      authenticating_with_token do
        ethon.http_request(
          uri_escape(
            "#{collaboration_base_uri}#{collaboration_site_path}/_api/web/GetFileByServerRelativeUrl('#{server_relative_url}')/ListItemAllFields/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)"
          ),
          :post,
          body: ''
        )
        ethon.perform
      end
    end

    # Resets role inheritance for a document
    def collaboration_reset_role_inheritance(server_relative_url)
      ethon = ethon_easy_json_requester
      ethon.headers = with_bearer_authentication_header({
                                                          'accept' => 'application/json;odata=verbose',
                                                          'X-RequestDigest' => xrequest_digest(collaboration_site_path)
                                                        })

      authenticating_with_token do
        ethon.http_request(
          uri_escape("#{collaboration_base_uri}#{collaboration_site_path}/_api/web/GetFileByServerRelativeUrl('#{server_relative_url}')/ListItemAllFields/resetroleinheritance"),
          :post,
          body: ''
        )
        ethon.perform
      end
    end

    # Grants permission to a specific user
    def collaboration_grant_user_permission(server_relative_url, user_id, role_id)
      ethon = ethon_easy_json_requester
      ethon.headers = with_bearer_authentication_header({
                                                          'accept' => 'application/json;odata=verbose',
                                                          'X-RequestDigest' => xrequest_digest(collaboration_site_path)
                                                        })

      authenticating_with_token do
        ethon.http_request(
          uri_escape(
            "#{collaboration_base_uri}#{collaboration_site_path}/_api/web/GetFileByServerRelativeUrl('#{server_relative_url}')/ListItemAllFields/roleassignments/addroleassignment(principalid=#{user_id},roleDefId=#{role_id})"
          ),
          :post,
          body: ''
        )
        ethon.perform
      end
    end

    # Revokes permission from a specific user
    def collaboration_revoke_permission(server_relative_url, user_id)
      ethon = ethon_easy_json_requester
      ethon.headers = with_bearer_authentication_header({
                                                          'accept' => 'application/json;odata=verbose',
                                                          'X-RequestDigest' => xrequest_digest(collaboration_site_path),
                                                          'IF-MATCH' => '*',
                                                          'X-HTTP-Method' => 'DELETE'
                                                        })

      authenticating_with_token do
        ethon.http_request(
          uri_escape(
            "#{collaboration_base_uri}#{collaboration_site_path}/_api/web/GetFileByServerRelativeUrl('#{server_relative_url}')/ListItemAllFields/roleassignments/getbyprincipalid(#{user_id})"
          ),
          :post,
          body: ''
        )
        ethon.perform
      end
    end

    def process_url(url, fields)
      easy = ethon_easy_json_requester
      easy.url = url
      easy.perform

      parsed_response_body = JSON.parse(easy.response_body)

      page_content = if fields
                       parsed_response_body['d']['results'].map { |v| v.fetch_values(*fields) }
                     else
                       parsed_response_body['d']['results']
                     end

      if next_url = parsed_response_body['d']['__next']
        page_content + process_url(next_url, fields)
      else
        page_content
      end
    end

    def token_auth?
      config.authentication == 'token'
    end

    def ntlm_auth?
      config.authentication == 'ntlm'
    end

    def with_bearer_authentication_header(h)
      return h if ntlm_auth?

      h.merge(bearer_auth_header)
    end

    def bearer_auth_header
      { 'Authorization' => bearer_auth }
    end

    def base_url
      config.uri
    end

    def base_api_url
      "#{base_url}/_api/"
    end

    def base_api_web_url
      "#{base_api_url}web/"
    end

    def computed_api_url(site)
      if site.present?
        "#{base_url}/#{site}/_api/"
      else
        "#{base_url}/_api/"
      end
    end

    def computed_web_api_url(site)
      remove_double_slashes("#{computed_api_url(site)}/web/")
    end

    def ethon_easy_json_requester
      easy = ethon_easy_requester
      easy.headers = with_bearer_authentication_header({ 'accept' => 'application/json;odata=verbose' })
      easy
    end

    def ethon_easy_options
      config.ethon_easy_options || {}
    end

    def ethon_easy_requester
      if token_auth?
        authenticating_with_token do
          easy = Ethon::Easy.new(DEFAULT_TOKEN_ETHON_OPTIONS.merge(ethon_easy_options))
          easy.headers = with_bearer_authentication_header({})
          easy
        end
      elsif ntlm_auth?
        easy = Ethon::Easy.new(DEFAULT_NTLM_ETHON_OPTIONS.merge(ethon_easy_options))
        easy.username = config.username
        easy.password = config.password
        easy
      end
    end

    # When you send a POST request, the request must include the form digest
    # value in the X-RequestDigest header
    def xrequest_digest(site_path = nil)
      easy = ethon_easy_json_requester
      url = remove_double_slashes("#{computed_api_url(site_path)}/contextinfo")
      easy.http_request(url, :post, { body: '' })
      easy.perform
      JSON.parse(easy.response_body)['d']['GetContextWebInformation']['FormDigestValue']
    end

    def last_location_header(ethon)
      last_redirect_idx = ethon.response_headers.rindex('HTTP/1.1 302')
      return if last_redirect_idx.nil?

      last_response_headers = ethon.response_headers[last_redirect_idx..]
      location = last_response_headers.match(/\r\n(Location:)(.+)\r\n/)[2].strip
      utf8_encode uri_unescape(location)
    end

    def check_and_raise_failure(ethon)
      return if (200..299).cover? ethon.response_code

      raise "Request failed, received #{ethon.response_code}:\n#{ethon.url}\n#{ethon.response_body}"
    end

    def prepare_metadata(metadata, type)
      metadata.inject("{ '__metadata': { 'type': '#{type}' }") do |result, element|
        key = element[0]
        value = element[1]
        result += ", '#{json_escape_single_quote(key.to_s)}': '#{json_escape_single_quote(value.to_s)}'"
      end + ' }'
    end

    def json_escape_single_quote(s)
      s.gsub("'", %q(\\\'))
    end

    def odata_escape_single_quote(s)
      s.gsub("'", "''")
    end

    def split_path(file_path)
      last_slash_pos = file_path.rindex('/')
      {
        path: file_path[0...last_slash_pos],
        name: file_path[(last_slash_pos + 1)..]
      }
    end

    def extract_paths(url)
      unescaped_url = string_unescape(url)
      uri = URI(uri_escape(unescaped_url))
      path = utf8_encode uri_unescape(uri.path)
      sites_match = /\/sites\/[^\/]+/.match(path)
      site_path = sites_match[0] unless sites_match.nil?
      file_path = site_path.nil? ? path : path.sub(site_path, '')
      uri.path = ''
      root_url = uri.to_s
      {
        root: root_url,
        site: site_path,
        file: file_path
      }
    end

    def validate_token_config
      valid_config_options(VALID_TOKEN_CONFIG_OPTIONS)
    end

    def validate_ntlm_config
      valid_config_options(VALID_NTLM_CONFIG_OPTIONS)
    end

    def valid_config_options(options = [])
      options.filter_map do |opt|
        c = config.send(opt)

        next if c.present? && string_not_blank?(c)

        opt
      end
    end

    def validate_config!
      raise Errors::InvalidAuthenticationError.new unless valid_authentication?(config.authentication)

      validate_token_config! if config.authentication == 'token'
      validate_ntlm_config! if config.authentication == 'ntlm'

      raise Errors::UriConfigurationError.new                       unless valid_uri?(config.uri)
      raise Errors::EthonOptionsConfigurationError.new              unless ethon_easy_options.is_a?(Hash)
    end

    def validate_token_config!
      invalid_token_opts = validate_token_config

      raise Errors::InvalidTokenConfigError.new(invalid_token_opts) unless invalid_token_opts.empty?
    end

    def validate_ntlm_config!
      invalid_ntlm_opts = validate_ntlm_config

      raise Errors::InvalidNTLMConfigError.new(invalid_ntlm_opts) unless invalid_ntlm_opts.empty?
    end

    def string_not_blank?(object)
      !object.nil? && object != '' && object.is_a?(String)
    end

    def valid_uri?(which)
      if which and which.is_a? String
        uri = URI.parse(which)
        uri.is_a?(URI::HTTP) || uri.is_a?(URI::HTTPS)
      else
        false
      end
    end

    def valid_authentication?(which)
      %w[ntlm token].include?(which)
    end

    # Waiting for RFC 3986 to be implemented, we need to escape square brackets
    def uri_escape(uri)
      URI_PARSER.escape(uri).gsub('[', '%5B').gsub(']', '%5D')
    end

    def uri_unescape(uri)
      URI_PARSER.unescape(uri.gsub('%5B', '[').gsub('%5D', ']'))
    end

    # TODO: Try to remove `eval` from this method. Otherwise, fix offenses
    # rubocop:disable Security/Eval, Style/DocumentDynamicEvalDefinition, Style/EvalWithLocation, Style/PercentLiteralDelimiters
    def string_unescape(s)
      s.gsub!(/\\(?:[abfnrtv])/, '') # remove control chars
      s.gsub!('"', '\"') # escape double quotes
      eval %Q{"#{s}"}
    end
    # rubocop:enable Security/Eval, Style/DocumentDynamicEvalDefinition, Style/EvalWithLocation, Style/PercentLiteralDelimiters

    def utf8_encode(s)
      s.force_encoding('UTF-8') unless s.nil?
    end

    def sanitize_filename(filename)
      escaped = Regexp.escape(FILENAME_INVALID_CHARS)
      regexp = Regexp.new("[#{escaped}]")
      sanitized_filename = filename.gsub(regexp, '-')
      if sanitized_filename.length > 128
        dot_index = sanitized_filename.rindex('.')
        if dot_index.nil?
          sanitized_filename = sanitized_filename[0..127]
        else
          extension_length = sanitized_filename.length - dot_index
          upper_bound = 127 - extension_length
          sanitized_filename = sanitized_filename[0..upper_bound] + sanitized_filename[dot_index...sanitized_filename.length]
        end
      end
      odata_escape_single_quote(sanitized_filename)
    end

    def build_search_kql_conditions(options)
      conditions = []
      conditions << 'IsContainer<>true'
      conditions << 'contentclass:STS_ListItem'
      conditions << "WebId=#{options[:web_id]}" unless options[:web_id].nil?
      conditions << "ListId:#{options[:list_id]}" unless options[:list_id].nil?
      "'#{conditions.join('+')}'"
    end

    def build_search_fql_conditions(options)
      start_at = options[:start_at]
      end_at = options[:end_at]
      if end_at.nil?
        "'write:range(#{start_at.utc.iso8601},max,from=\"ge\")'"
      else
        "'write:range(#{start_at.utc.iso8601},#{end_at.utc.iso8601},from=\"ge\",to=\"le\")'"
      end
    end

    def build_search_properties(options)
      default_properties = %w[
        Write IsDocument IsContainer ListId WebId URL
        Created Title Author Size Path UniqueId contentclass
      ]
      properties = options[:properties] || []
      properties += default_properties
      "selectproperties='#{properties.join(',')}'"
    end

    def build_search_paging(options)
      start = options[:start_result] || 0
      max = options[:max_results] || 500
      "startrow=#{start}&rowlimit=#{max}"
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

    def parse_list_response(json_response, all_properties)
      results = json_response['d']['results']
      records = []
      results.each do |result|
        # Skip folders
        next unless result['FileSystemObjectType'].eql? 0

        record = {}
        (all_properties - %w[File URL]).each do |key|
          record[key.underscore.to_sym] = result[key]
        end
        file = result['File']
        %w[Name ServerRelativeUrl Length].each do |key|
          record[key.underscore.to_sym] = file[key]
        end
        record[:url] = result['URL'].nil? ? nil : result['URL']['Url']
        records << OpenStruct.new(record)
      end
      records
    end

    def parse_get_document_response(response_body, custom_properties)
      all_props = JSON.parse(response_body)['d']
      default_properties = %w[GUID Title Created Modified]
      keys = default_properties + custom_properties
      props = {}
      keys.each do |key|
        props[key.underscore.to_sym] = all_props[key]
      end
      props[:url] = all_props['URL'].nil? ? nil : all_props['URL']['Url']
      OpenStruct.new(props)
    end

    def remove_double_slashes(str)
      str.to_s.gsub('//', '/')
              .gsub('http:/', 'http://')
              .gsub('https:/', 'https://')
    end

    def build_query_params(query)
      query_params = []

      query.each do |field, value|
        query_params << "$#{field}=#{value}"
      end

      query_params.join('&')
    end

    def parse_lists_in_site_response(response_body)
      json_response = JSON.parse(response_body)
      results = json_response.dig('d', 'results')

      results.map do |result|
        OpenStruct.new(result.transform_keys { |k| k.underscore.to_sym })
      end
    end

    def update_object_metadata(metadata, new_metadata, site_path = '')
      update_metadata_url = metadata['uri']
      prepared_metadata = prepare_metadata(new_metadata, metadata['type'])

      easy = ethon_easy_json_requester
      easy.headers = with_bearer_authentication_header({ 'accept' => 'application/json;odata=verbose',
                                                         'content-type' => 'application/json;odata=verbose',
                                                         'X-RequestDigest' => xrequest_digest(site_path),
                                                         'X-Http-Method' => 'PATCH',
                                                         'If-Match' => '*' })

      easy.http_request(update_metadata_url,
                        :post,
                        { body: prepared_metadata })
      easy.perform
      check_and_raise_failure(easy)
      easy.response_code
    end
  end
end
