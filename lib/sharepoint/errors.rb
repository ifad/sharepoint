module Sharepoint
  module Errors
    class UriConfigurationError < StandardError
      def initialize
        super('Invalid Uri configuration')
      end
    end

    class EthonOptionsConfigurationError < StandardError
      def initialize
        super('Invalid ethon easy options')
      end
    end

    class InvalidOauthConfigError < StandardError
      def initialize(invalid_entries)
        error_messages = invalid_entries.map do |e|
          "Invalid #{e} in OAUTH configuration"
        end

        super error_messages.join(',')
      end
    end
  end
end
