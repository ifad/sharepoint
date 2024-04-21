require 'ostruct'
require 'ethon'
require 'uri'
require 'json'
require 'time'

require 'active_support/core_ext/string/inflections'
require 'active_support/core_ext/object/blank'

module Sharepoint
  class Client
    FILENAME_INVALID_CHARS = '~"#%&*:<>?/\{|}'

    # @return [OpenStruct] The current configuration.
    attr_reader :config

    # Initializes a new client with given options.
    #
    # @param [Hash] config The client options:
    #  - `:uri` The SharePoint server's root url
    #  - `:username` self-explanatory
    #  - `:password` self-explanatory
    # @return [Sharepoint::Client] client object
    def initialize(config = {})
      @config = OpenStruct.new(config)
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
      if meta.url.nil?
        url = computed_web_api_url(site_path)
        server_relative_url = odata_escape_single_quote "#{site_path}#{file_path}"
        download_url "#{url}GetFileByServerRelativeUrl('#{server_relative_url}')/$value"
      else # requested file is a link
        paths = extract_paths(meta.url)
        link_config = { uri: paths[:root] }
        if link_credentials.empty?
          link_config = config.to_h.merge(link_config)
        else
          link_config.merge!(link_credentials)
        end
        link_client = self.class.new(link_config)
        result = link_client.download_url meta.url
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
      path = path[1..-1] if path[0].eql?('/')
      url = uri_escape "#{url}GetFolderByServerRelativeUrl('#{path}')/Folders"
      easy = ethon_easy_json_requester
      easy.headers = { 'accept' => 'application/json;odata=verbose',
                       'content-type' => 'application/json;odata=verbose',
                       'X-RequestDigest' => xrequest_digest(site_path) }
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
      path = path[1..-1] if path[0].eql?('/')
      url = uri_escape "#{url}GetFolderByServerRelativeUrl('#{path}')/Files/Add(url='#{sanitized_filename}',overwrite=true)"
      easy = ethon_easy_json_requester
      easy.headers = { 'accept' => 'application/json;odata=verbose',
                       'X-RequestDigest' => xrequest_digest(site_path) }
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
      easy.headers = { 'accept' =>  'application/json;odata=verbose',
                       'content-type' =>  'application/json;odata=verbose',
                       'X-RequestDigest' =>  xrequest_digest(site_path),
                       'X-Http-Method' =>  'PATCH',
                       'If-Match' => '*' }
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
      url = "#{computed_web_api_url(site_path)}Lists".dup
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

    private

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
      easy.headers = { 'accept' => 'application/json;odata=verbose' }
      easy
    end

    def ethon_easy_options
      config.ethon_easy_options || {}
    end

    def ethon_easy_requester
      easy = Ethon::Easy.new({ httpauth: :ntlm, followlocation: 1, maxredirs: 5 }.merge(ethon_easy_options))
      easy.username = config.username
      easy.password = config.password
      easy
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

      last_response_headers = ethon.response_headers[last_redirect_idx..-1]
      location = last_response_headers.match(/\r\n(Location:)(.+)\r\n/)[2].strip
      utf8_encode uri_unescape(location)
    end

    def check_and_raise_failure(ethon)
      return if (200..299).cover? ethon.response_code

      raise "Request failed, received #{ethon.response_code}:\n#{ethon.url}\n#{ethon.response_body}"
    end

    def prepare_metadata(metadata, type)
      metadata.inject("{ '__metadata': { 'type': '#{type}' }") { |result, element|
        key = element[0]
        value = element[1]
        result += ", '#{json_escape_single_quote(key.to_s)}': '#{json_escape_single_quote(value.to_s)}'"
      } + ' }'
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
        path: file_path[0..last_slash_pos - 1],
        name: file_path[last_slash_pos + 1..-1]
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

    def validate_config!
      raise Errors::UsernameConfigurationError.new      unless string_not_blank?(@config.username)
      raise Errors::PasswordConfigurationError.new      unless string_not_blank?(@config.password)
      raise Errors::UriConfigurationError.new           unless valid_config_uri?
      raise Errors::EthonOptionsConfigurationError.new  unless ethon_easy_options.is_a?(Hash)
    end

    def string_not_blank?(object)
      !object.nil? && object != '' && object.is_a?(String)
    end

    def valid_config_uri?
      if @config.uri and @config.uri.is_a? String
        uri = URI.parse(@config.uri)
        uri.is_a?(URI::HTTP) || uri.is_a?(URI::HTTPS)
      else
        false
      end
    end

    # Waiting for RFC 3986 to be implemented, we need to escape square brackets
    def uri_escape(uri)
      URI::DEFAULT_PARSER.escape(uri).gsub('[', '%5B').gsub(']', '%5D')
    end

    def uri_unescape(uri)
      URI::DEFAULT_PARSER.unescape(uri.gsub('%5B', '[').gsub('%5D', ']'))
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
          sanitized_filename = sanitized_filename[0..upper_bound] + sanitized_filename[dot_index..sanitized_filename.length - 1]
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
        OpenStruct.new(result.map { |k, v| [k.underscore.to_sym, v] }.to_h)
      end
    end

    def update_object_metadata(metadata, new_metadata, site_path = '')
      update_metadata_url = metadata['uri']
      prepared_metadata = prepare_metadata(new_metadata, metadata['type'])

      easy = ethon_easy_json_requester
      easy.headers = { 'accept' =>  'application/json;odata=verbose',
                       'content-type' =>  'application/json;odata=verbose',
                       'X-RequestDigest' =>  xrequest_digest(site_path),
                       'X-Http-Method' =>  'PATCH',
                       'If-Match' => '*' }

      easy.http_request(update_metadata_url,
                        :post,
                        { body: prepared_metadata })
      easy.perform
      check_and_raise_failure(easy)
      easy.response_code
    end
  end
end
