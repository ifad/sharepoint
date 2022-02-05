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

      it "sets base_url in the client" do
        expect(subject.send :base_url).to eql(ENV['SP_URL'])
      end

      it "sets base_api_url in the client" do
        expect(subject.send :base_api_url).to eql(ENV['SP_URL']+'/_api/')
      end

      it "sets base_api_web_url in the client" do
        expect(subject.send :base_api_web_url).to eql(ENV['SP_URL']+'/_api/web/')
      end
    end

    context 'ethon easy options' do
      context 'success' do
        let(:config_ethon) { config.merge({ ethon_easy_options: ssl_verifypeer }) }
        let(:ssl_verifypeer) { { ssl_verifypeer: false } }

        subject { described_class.new(config_ethon) }

        it "sets ethon easy options in the client" do
          expect(subject.send :ethon_easy_options).to eql(ssl_verifypeer)
        end
      end

      context 'failure' do
        let(:config_ethon) { config.merge({ ethon_easy_options: 'hello' }) }

        it "should raise ethon configuration error for bad config" do
          expect {
            described_class.new(config_ethon)
          }.to raise_error(Sharepoint::Errors::EthonOptionsConfigurationError)
        end
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

  describe '#remove_double_slashes' do
    PAIRS = {
      'foobar'             => 'foobar',
      'foo/bar'            => 'foo/bar',
      'foo/bar/'           => 'foo/bar/',
      'http://foo/bar//'   => 'http://foo/bar/',
      'https://foo/bar//'  => 'https://foo/bar/',
      'https://foo/bar'    => 'https://foo/bar',
      'https://foo//bar//' => 'https://foo/bar/'
    }.each do |input, output|
      specify do
        expect(described_class.new(config).send :remove_double_slashes, input).to eq(output)
      end
    end
  end

  {
    '[]'             => '%5B%5D',
    "https://example.org/sites/Method('/file+name .pdf')" => "https://example.org/sites/Method('/file+name%20.pdf')"
  }.each do |input, output|
    describe '#uri_escape' do
      specify do
        expect(described_class.new(config).send :uri_escape, input).to eq(output)
      end
    end

    describe '#uri_unescape' do
      specify do
        expect(described_class.new(config).send :uri_unescape, output).to eq(input)
      end
    end
  end

end
