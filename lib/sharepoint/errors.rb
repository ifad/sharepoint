module Sharepoint
  module Errors
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
    class EthonOptionsConfigurationError < StandardError
      def initialize
        super "Invalid ethon easy options"
      end
    end
  end
end
