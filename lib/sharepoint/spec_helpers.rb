module Sharepoint
  module SpecHelpers
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

    def mock_requests
      allow_any_instance_of(Ethon::Easy)
        .to receive(:perform)
        .and_return(nil)
    end

    def mock_responses(fixture_file)
      allow_any_instance_of(Ethon::Easy)
        .to receive(:response_code)
        .and_return(200)
      allow_any_instance_of(Ethon::Easy)
        .to receive(:response_body)
        .and_return(
          File.open("spec/fixtures/responses/#{fixture_file}").read
        )
    end
  end
end
