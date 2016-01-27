module Helpers
  def value_to_string(value)
    case value
      when nil
        "nil"
      when ""
        "blank"
      else
        value
    end
  end
end

RSpec.configure do |config|
  config.include Helpers
end
