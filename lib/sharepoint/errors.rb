module Sharepoint
  module Errors
    class ClientNotInitialized < StandardError
      def initialize
        super "Default client and cofiguration not initialized"
      end
    end
    class InvalidClient < StandardError
      def initialize
        super "Assigned client is not a Sharepoint::Client"
      end
    end
    class UsernameConfigurationError < StandardError
      def initialize
        super "Invalid Username Configuration"
      end
    end
    class PasswordConfigurationError < StandardError
      def initialize
        super "Invalid Password configuration"
      end
    end
    class UriConfigurationError < StandardError
      def initialize
        super "Invalid Uri configuration"
      end
    end
  end
end
