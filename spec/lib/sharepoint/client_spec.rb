# frozen_string_literal: true

require 'spec_helper'

describe Sharepoint::Client do
  before { ENV['SP_URL'] = 'https://localhost:8888' }

  let(:config) { sp_config }

  describe '#initialize' do
    context 'with success' do
      subject(:client) { described_class.new(config) }

      it 'returns a valid instance' do
        expect(client).to be_a described_class
      end

      it 'sets config object' do
        client_config = client.config
        expect(client_config).to be_a OpenStruct
        [:client_id, :client_secret, :tenant_id, :cert_name, :auth_scope, :url].each do |key|
          value = client_config.send(key)
          expect(value).to eq config[key]
        end
      end

      it 'sets base_url in the client' do
        expect(client.send(:base_url)).to eql(ENV.fetch('SP_URL', nil))
      end

      it "sets base_api_url in the client" do
        expect(subject.send :base_api_url).to eql("#{ENV['SP_URL']}/_api/")
      end

      it "sets base_api_web_url in the client" do
        expect(subject.send :base_api_web_url).to eql("#{ENV['SP_URL']}/_api/web/")
      end
    end

    context "correct authentication" do
      [{ value: "ntlm",  name:   'ntlm' },
       { value: "token", name:   'token' } ].each do |ocurrence|

        it "should not raise authentication configuration error for #{ ocurrence[:name]} authentication" do
          correct_config = config
          correct_config[:authentication] = ocurrence[:value]

          expect {
           described_class.new(correct_config)
          }.not_to raise_error(Sharepoint::Errors::InvalidAuthenticationError)
        end
      end
    end

    context 'ethon easy options' do
      context 'success' do
        let(:config_ethon) { config.merge({ ethon_easy_options: ssl_verifypeer }) }
        let(:ssl_verifypeer) { { ssl_verifypeer: false } }

        it 'sets ethon easy options in the client' do
          expect(described_class.new(config_ethon).send(:ethon_easy_options)).to eql(ssl_verifypeer)
        end
      end

      context 'with failure' do
        let(:config_ethon) { config.merge({ ethon_easy_options: 'hello' }) }

        it 'raises ethon configuration error for bad config' do
          expect do
            described_class.new(config_ethon)
          end.to raise_error(Sharepoint::Errors::EthonOptionsConfigurationError)
        end
      end
    end

    context 'failure' do

      context "bad authentication" do
        [{ value: nil, name:   'nil' },
         { value:  '', name: 'blank' },
         { value: 344, name:     344 } ].each do |ocurrence|

          it "should raise authentication configuration error for #{ ocurrence[:name]} authentication" do
            wrong_config = config
            wrong_config[:authentication] = ocurrence[:value]

            expect {
             described_class.new(wrong_config)
            }.to raise_error(Sharepoint::Errors::InvalidAuthenticationError)
          end
        end
      end
      
      context "token" do
        before { ENV['SP_AUTHENTICATION'] = 'token' }

        context "bad client_id" do
          [{ value: nil, name:   'nil' },
          { value:  '', name: 'blank' },
          { value: 344, name:     344 } ].each do |ocurrence|

            it "should raise client_id configuration error for #{ ocurrence[:name]} client_id" do
              wrong_config = config
              wrong_config[:client_id] = ocurrence[:value]

              expect {
                described_class.new(wrong_config)
              }.to raise_error(Sharepoint::Errors::InvalidTokenConfigError)
            end
          end
        end

        context "bad client_secret" do
          [{ value: nil, name:   'nil' },
          { value:  '', name: 'blank' },
          { value: 344, name:     344 } ].each do |ocurrence|

            it "should raise client_secret configuration error for #{ocurrence[:name]} client_secret" do
              wrong_config = config
              wrong_config[:client_secret] = ocurrence[:value]

              expect {
                described_class.new(wrong_config)
              }.to raise_error(Sharepoint::Errors::InvalidTokenConfigError)
            end
          end
        end

        context "bad tenant_id" do
          [{ value: nil, name:   'nil' },
          { value:  '', name: 'blank' },
          { value: 344, name:     344 } ].each do |ocurrence|

            it "should raise tenant_id configuration error for #{ocurrence[:name]} tenant_id" do
              wrong_config = config
              wrong_config[:tenant_id] = ocurrence[:value]

              expect {
                described_class.new(wrong_config)
              }.to raise_error(Sharepoint::Errors::InvalidTokenConfigError)
            end
          end
        end

        context "bad cert_name" do
          [{ value: nil, name:   'nil' },
          { value:  '', name: 'blank' },
          { value: 344, name:     344 } ].each do |ocurrence|

            it "should raise cert_name configuration error for #{ocurrence[:name]} cert_name" do
              wrong_config = config
              wrong_config[:cert_name] = ocurrence[:value]

              expect {
                described_class.new(wrong_config)
              }.to raise_error(Sharepoint::Errors::InvalidTokenConfigError)
            end
          end
        end

        context "bad auth_scope" do
          [{ value:                  nil, name:         'nil' },
          { value:                   '', name:       'blank' },
          { value:                  344, name:           344 }].each do |ocurrence|

            it "should raise auth_scope configuration error for #{ocurrence[:name]} auth_scope" do
              wrong_config = config
              wrong_config[:auth_scope] = ocurrence[:value]

              expect {
                described_class.new(wrong_config)
              }.to raise_error(Sharepoint::Errors::InvalidTokenConfigError)
            end
          end

        end

        context "bad auth_scope" do
          [{ value: 'ftp://www.test.com', name: "invalid auth_scope" }].each do |ocurrence|

            it "should raise auth_scope configuration error for #{ocurrence[:name]} auth_scope" do
              wrong_config = config
              wrong_config[:auth_scope] = ocurrence[:value]

              expect {
                described_class.new(wrong_config)
              }.to raise_error(Sharepoint::Errors::UriConfigurationError)
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

      context "ntlm" do
        before { ENV['SP_AUTHENTICATION'] = 'ntlm' }

        context "bad username" do
          [{ value: nil, name:   'nil' },
          { value:  '', name: 'blank' },
          { value: 344, name:     344 } ].each do |ocurrence|

            it "should raise username configuration error for #{ ocurrence[:name]} username" do
              wrong_config = config
              wrong_config[:username] = ocurrence[:value]

              expect {
              described_class.new(wrong_config)
              }.to raise_error(Sharepoint::Errors::InvalidNTLMConfigError)
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
              }.to raise_error(Sharepoint::Errors::InvalidNTLMConfigError)
            end
          end
        end
      end
    end
  end

  describe '#remove_double_slashes' do
    {
      'foobar' => 'foobar',
      'foo/bar' => 'foo/bar',
      'foo/bar/' => 'foo/bar/',
      'http://foo/bar//' => 'http://foo/bar/',
      'https://foo/bar//' => 'https://foo/bar/',
      'https://foo/bar' => 'https://foo/bar',
      'https://foo//bar//' => 'https://foo/bar/'
    }.each do |input, output|
      specify do
        expect(described_class.new(config).send(:remove_double_slashes, input)).to eq(output)
      end
    end
  end

  {
    '[]' => '%5B%5D',
    "https://example.org/sites/Method('/file+name .pdf')" => "https://example.org/sites/Method('/file+name%20.pdf')"
  }.each do |input, output|
    describe '#uri_escape' do
      specify do
        expect(described_class.new(config).send(:uri_escape, input)).to eq(output)
      end
    end

    describe '#uri_unescape' do
      specify do
        expect(described_class.new(config).send(:uri_unescape, output)).to eq(input)
      end
    end
  end
end
