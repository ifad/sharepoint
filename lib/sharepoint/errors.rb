module Sharepoint
  module Errors
    class UsernameConfigurationError < StandardError
      def initialize
        super('Invalid Username Configuration')
      end
    end

    class PasswordConfigurationError < StandardError
      def initialize
        super('Invalid Password configuration')
      end
    end

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
