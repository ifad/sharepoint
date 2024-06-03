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

    class InvalidAuthenticationError < StandardError
      def initialize
        super('Invalid authentication mechanism')
      end
    end

    class InvalidTokenConfigError < StandardError
      def initialize(invalid_entries)
        error_messages = invalid_entries.map do |e|
          "Invalid #{e} in Token configuration"
        end

        super(error_messages.join(','))
      end
    end

    class InvalidNTLMConfigError < StandardError
      def initialize(invalid_entries)
        error_messages = invalid_entries.map do |e|
          "Invalid #{e} in NTLM configuration"
        end

        super(error_messages.join(','))
      end
    end
  end
end
