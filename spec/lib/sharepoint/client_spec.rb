require 'spec_helper'

describe Sharepoint::Client do
  let(:config)    { { username: ENV['SP_USERNAME'],
                      password: ENV['SP_PASSWORD'],
                      uri:      ENV['SP_URL'] } }
  let(:client)    { described_class.new(config) }

  context "client undefined" do
    describe ".client" do
      it "should raise client undefined error" do
        expect {
          described_class.client
        }.to raise_error(Sharepoint::Errors::ClientNotInitialized)
      end
    end

    describe ".client=" do
      it "should raise invalid client error" do
        expect {
          described_class.client = "Invalid Client"
        }.to raise_error(Sharepoint::Errors::InvalidClient)
      end
    end

    describe ".config" do
      it "should raise client undefined error" do
        expect {
          described_class.config
        }.to raise_error(Sharepoint::Errors::ClientNotInitialized)
      end
    end

    describe ".initialize" do
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

    describe ".documents_for" do
      it "shoud raise client undefined error" do
        expect {
          described_class.documents_for("path")
        }.to raise_error(Sharepoint::Errors::ClientNotInitialized)
      end
    end

    describe ".upload" do
      it "shoud raise client undefined error" do
        expect {
          described_class.upload("filename", "content", "path")
        }.to raise_error(Sharepoint::Errors::ClientNotInitialized)
      end
    end

    describe ".update_metadata" do
      it "shoud raise client undefined error" do
        expect {
          described_class.update_metadata("filename", { key1: "metadata1" },"path")
        }.to raise_error(Sharepoint::Errors::ClientNotInitialized)
      end
    end
  end

  context "client defined" do
    before :each do
      described_class.client = client
    end

    describe ".client .client=" do
      it "return the default client" do
        expect(described_class.client).to eq(client)
      end
    end

    describe ".config" do
      it "raise client undefined error" do
        expect(described_class.config).to eq(OpenStruct.new(config))
      end
    end

    describe ".initialize" do
      it "define @user instance var in the client" do
        expect(described_class.client
                              .instance_variable_get(:@user)).to eql(ENV['SP_USERNAME'])
      end

      it "define @password instance var in the client" do
        expect(described_class.client
                              .instance_variable_get(:@password)).to eql(ENV['SP_PASSWORD'])
      end

      it "define @base_url instance var in the client" do
        expect(described_class.client
                              .instance_variable_get(:@base_url)).to eql(ENV['SP_URL'])
      end

      it "define @base_api_url instance var in the client" do
        expect(described_class.client
                              .instance_variable_get(:@base_api_url)).to eql(ENV['SP_URL']+'/_api/web/')
      end
    end

    describe ".upload" do
      described_class::FILENAME_INVALID_CHARS.each do |char|
        it "shoud raise invalid file name error if the filename contains the character " + char do
          expect {
            described_class.upload(char + "filename", "content", "path")
          }.to raise_error(Sharepoint::Errors::InvalidSharepointFilename)
        end
      end
      it "should upload the file correctly"
    end

    describe ".metadata" do
      it "shoud raise invalid metadata if any metadata value or key include the single quote char"
      it "should update the metadata correctly"
    end
  end
end
