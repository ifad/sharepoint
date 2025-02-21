# frozen_string_literal: true

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

    def sp_config(authentication: nil)
      {
        uri: ENV.fetch('SP_URL', nil),
        authentication: authentication || ENV.fetch('SP_AUTHENTICATION', nil),
        client_id: ENV.fetch('SP_CLIENT_ID', nil),
        client_secret: ENV.fetch('SP_CLIENT_SECRET', nil),
        tenant_id: ENV.fetch('SP_TENANT_ID', nil),
        cert_name: ENV.fetch('SP_CERT_NAME', nil),
        auth_scope: ENV.fetch('SP_AUTH_SCOPE', nil),
        username: ENV.fetch('SP_USERNAME', nil),
        password: ENV.fetch('SP_PASSWORD', nil)
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
        .and_return({ 'Token' => { 'expires_in' => 3600, 'access_token' => 'access_token' } })
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
