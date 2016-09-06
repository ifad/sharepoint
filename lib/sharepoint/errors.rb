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

    class InvalidSharepointFilename < StandardError
      def initialize
        super "The file name contains an invalid character"
      end
    end

  end
end
