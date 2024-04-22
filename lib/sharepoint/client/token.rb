module Sharepoint
  class Client
    class Token
      class InvalidTokenError < StandardError
      end

      attr_accessor :expires_in, :access_token, :fetched_at
      attr_reader :config

      def initialize(config)
        @config = config
      end

      def get_or_fetch
        return access_token unless access_token.nil? || expired?

        fetch
      end

      def to_s
        access_token
      end

      def fetch
        response = request_new_token

        details = response['Token']
        self.fetched_at = Time.now.utc.to_i
        self.expires_in = details['expires_in']
        self.access_token = details['access_token']
      end

      private

      def expired?
        return true unless fetched_at && expires_in

        (fetched_at + expires_in) < Time.now.utc.to_i
      end

      def
      def(_request_new_token)
        auth_request = {
          client_id: config.client_id,
          client_secret: config.client_secret,
          tenant_id: config.tenant_id,
          cert_name: config.cert_name,
          auth_scope: config.auth_scope
        }.to_json

        headers = { 'Content-Type' => 'application/json' }

        ethon = Ethon::Easy.new(followlocation: true)
        ethon.http_request(config.token_url, :post, body: auth_request, headers: headers)
        ethon.perform

        raise InvalidTokenError.new(ethon.response_body.to_s) unless ethon.response_code == 200

        JSON.parse(ethon.response_body)
      end
    end
  end
end
