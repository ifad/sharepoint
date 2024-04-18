module Sharepoint
  module SpecHelpers
    def value_to_string(value)
      case value
      when nil
        'nil'
      when ''
        'blank'
      else
        value
      end
    end

    def sp_config
      {
        uri: ENV['SP_URL'],
        authentication: ENV['SP_AUTHENTICATION'],
        client_id: ENV['SP_CLIENT_ID'],
        client_secret: ENV['SP_CLIENT_SECRET'],
        tenant_id: ENV['SP_TENANT_ID'],
        cert_name: ENV['SP_CERT_NAME'],
        auth_scope: ENV['SP_AUTH_SCOPE'],
        username: ENV['SP_USERNAME'],
        password: ENV['SP_PASSWORD']
      }
    end

    def mock_requests
      allow_any_instance_of(Ethon::Easy)
        .to receive(:perform)
        .and_return(nil)
    end

    def mock_token_responses
      allow_any_instance_of(Sharepoint::Client::Token)
        .to receive(:request_new_token)
        .and_return({"Token" => { "expires_in" => 3600, "access_token" => "access_token" }})
    end

    def mock_responses(fixture_file)
      allow_any_instance_of(Ethon::Easy)
        .to receive(:response_code)
        .and_return(200)
      allow_any_instance_of(Ethon::Easy)
        .to receive(:response_body)
        .and_return(
          File.read("spec/fixtures/responses/#{fixture_file}")
        )
    end
  end
end
