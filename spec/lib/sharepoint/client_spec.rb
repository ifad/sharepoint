require 'spec_helper'

describe Sharepoint::Client do
  let(:config)    { { username: ENV['SP_USERNAME'],
                      password: ENV['SP_PASSWORD'],
                      uri:      ENV['SP_URL'] } }

  describe '#initialize' do

    context 'success' do
      subject { described_class.new(config) }

      it 'returns a valid instance' do
        is_expected.to be_a Sharepoint::Client
      end

      it 'sets config object' do
        client_config = subject.config
        expect(client_config).to be_a OpenStruct
        [:username, :password, :url].each do |key|
          value = client_config.send(key)
          expect(value).to eq config[key]
        end
      end

      it "defines @user instance var in the client" do
        expect(
          subject.instance_variable_get(:@user)
        ).to eql(ENV['SP_USERNAME'])
      end

      it "defines @password instance var in the client" do
        expect(
          subject.instance_variable_get(:@password)
        ).to eql(ENV['SP_PASSWORD'])
      end

      it "defines @base_url instance var in the client" do
        expect(
          subject.instance_variable_get(:@base_url)
        ).to eql(ENV['SP_URL'])
      end

      it "defines @base_api_url instance var in the client" do
        expect(
          subject.instance_variable_get(:@base_api_url)
        ).to eql(ENV['SP_URL']+'/_api/')
      end

      it "defines @base_api_web_url instance var in the client" do
        expect(
          subject.instance_variable_get(:@base_api_web_url)
        ).to eql(ENV['SP_URL']+'/_api/web/')
      end
    end

    context 'failure' do

      context "bad username" do
        [{ value: nil, name:   'nil' },
         { value:  '', name: 'blank' },
         { value: 344, name:     344 } ].each do |ocurrence|

          it "should raise username configuration error for #{ ocurrence[:name]} username" do
            wrong_config = config
            wrong_config[:username] = ocurrence[:value]

            expect {
             described_class.new(wrong_config)
            }.to raise_error(Sharepoint::Errors::UsernameConfigurationError)
          end
        end
      end

      context "bad password" do
        [{ value: nil, name:   'nil' },
         { value:  '', name: 'blank' },
         { value: 344, name:     344 } ].each do |ocurrence|

          it "should raise password configuration error for #{ocurrence[:name]} password" do
            wrong_config = config
            wrong_config[:password] = ocurrence[:value]

            expect {
              described_class.new(wrong_config)
            }.to raise_error(Sharepoint::Errors::PasswordConfigurationError)
          end
        end
      end

      context "bad uri" do
        [{ value:                  nil, name:         'nil' },
         { value:                   '', name:       'blank' },
         { value:                  344, name:           344 },
         { value: 'ftp://www.test.com', name: "invalid uri" }].each do |ocurrence|

          it "should raise uri configuration error for #{ocurrence[:name]} uri" do
            wrong_config = config
            wrong_config[:uri] = ocurrence[:value]

            expect {
              described_class.new(wrong_config)
            }.to raise_error(Sharepoint::Errors::UriConfigurationError)
          end
        end

      end

    end

  end

end
