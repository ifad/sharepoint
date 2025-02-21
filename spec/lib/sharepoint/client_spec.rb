# frozen_string_literal: true

require 'spec_helper'

RSpec.describe Sharepoint::Client do
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
        %i[client_id client_secret tenant_id cert_name auth_scope url].each do |key|
          value = client_config.send(key)
          expect(value).to eq config[key]
        end
      end

      it 'sets base_url in the client' do
        expect(client.send(:base_url)).to eql(ENV.fetch('SP_URL', nil))
      end

      it 'sets base_api_url in the client' do
        expect(client.send(:base_api_url)).to eql("#{ENV.fetch('SP_URL', nil)}/_api/")
      end

      it 'sets base_api_web_url in the client' do
        expect(client.send(:base_api_web_url)).to eql("#{ENV.fetch('SP_URL', nil)}/_api/web/")
      end
    end

    context 'with authentication' do
      [{ value: 'ntlm',  name:   'ntlm' },
       { value: 'token', name:   'token' }].each do |occurrence|
        it "does not raise authentication configuration error for #{occurrence[:name]} authentication" do
          correct_config = config
          correct_config[:authentication] = occurrence[:value]

          expect do
            described_class.new(correct_config)
          end.not_to raise_error
        end
      end
    end

    context 'with ethon easy options' do
      context 'with success' do
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

    context 'with failure' do
      context 'with bad authentication' do
        [{ value: nil, name: 'nil' },
         { value: '', name: 'blank' },
         { value: 344, name: 344 }].each do |occurrence|
          it "raises authentication configuration error for #{occurrence[:name]} authentication" do
            wrong_config = config
            wrong_config[:authentication] = occurrence[:value]

            expect do
              described_class.new(wrong_config)
            end.to raise_error(Sharepoint::Errors::InvalidAuthenticationError)
          end
        end
      end

      context 'with token' do
        before { ENV['SP_AUTHENTICATION'] = 'token' }

        context 'with bad client_id' do
          [{ value: nil, name: 'nil' },
           { value: '', name: 'blank' },
           { value: 344, name: 344 }].each do |occurrence|
            it "raises client_id configuration error for #{occurrence[:name]} client_id" do
              wrong_config = config
              wrong_config[:client_id] = occurrence[:value]

              expect do
                described_class.new(wrong_config)
              end.to raise_error(Sharepoint::Errors::InvalidTokenConfigError)
            end
          end
        end

        context 'with bad client_secret' do
          [{ value: nil, name: 'nil' },
           { value: '', name: 'blank' },
           { value: 344, name: 344 }].each do |occurrence|
            it "raises client_secret configuration error for #{occurrence[:name]} client_secret" do
              wrong_config = config
              wrong_config[:client_secret] = occurrence[:value]

              expect do
                described_class.new(wrong_config)
              end.to raise_error(Sharepoint::Errors::InvalidTokenConfigError)
            end
          end
        end

        context 'with bad tenant_id' do
          [{ value: nil, name: 'nil' },
           { value: '', name: 'blank' },
           { value: 344, name: 344 }].each do |occurrence|
            it "raises tenant_id configuration error for #{occurrence[:name]} tenant_id" do
              wrong_config = config
              wrong_config[:tenant_id] = occurrence[:value]

              expect do
                described_class.new(wrong_config)
              end.to raise_error(Sharepoint::Errors::InvalidTokenConfigError)
            end
          end
        end

        context 'with bad cert_name' do
          [{ value: nil, name: 'nil' },
           { value: '', name: 'blank' },
           { value: 344, name: 344 }].each do |occurrence|
            it "raises cert_name configuration error for #{occurrence[:name]} cert_name" do
              wrong_config = config
              wrong_config[:cert_name] = occurrence[:value]

              expect do
                described_class.new(wrong_config)
              end.to raise_error(Sharepoint::Errors::InvalidTokenConfigError)
            end
          end
        end

        context 'with bad auth_scope' do
          [{ value: nil, name: 'nil' },
           { value: '', name: 'blank' },
           { value: 344, name: 344 }].each do |occurrence|
            it "raises auth_scope configuration error for #{occurrence[:name]} auth_scope" do
              wrong_config = config
              wrong_config[:auth_scope] = occurrence[:value]

              expect do
                described_class.new(wrong_config)
              end.to raise_error(Sharepoint::Errors::InvalidTokenConfigError)
            end
          end
        end

        it 'with bad auth_scope uri format' do
          skip 'Uri is not formatted'
          [{ value: 'ftp://www.test.com', name: 'invalid auth_scope' }].each do |occurrence|
            it "raises auth_scope configuration error for #{occurrence[:name]} auth_scope" do
              wrong_config = config
              wrong_config[:auth_scope] = occurrence[:value]

              expect do
                described_class.new(wrong_config)
              end.to raise_error(Sharepoint::Errors::UriConfigurationError)
            end
          end
        end

        context 'with bad uri' do
          [{ value: nil, name: 'nil' },
           { value: '', name: 'blank' },
           { value: 344, name: 344 },
           { value: 'ftp://www.test.com', name: 'invalid uri' }].each do |occurrence|
            it "raises uri configuration error for #{occurrence[:name]} uri" do
              wrong_config = config
              wrong_config[:uri] = occurrence[:value]

              expect do
                described_class.new(wrong_config)
              end.to raise_error(Sharepoint::Errors::UriConfigurationError)
            end
          end
        end
      end

      context 'when ntlm' do
        before { ENV['SP_AUTHENTICATION'] = 'ntlm' }

        context 'with bad username' do
          [{ value: nil, name: 'nil' },
           { value: '', name: 'blank' },
           { value: 344, name: 344 }].each do |occurrence|
            it "raises username configuration error for #{occurrence[:name]} username" do
              wrong_config = config
              wrong_config[:username] = occurrence[:value]

              expect do
                described_class.new(wrong_config)
              end.to raise_error(Sharepoint::Errors::InvalidNTLMConfigError)
            end
          end
        end

        context 'with bad password' do
          [{ value: nil, name: 'nil' },
           { value: '', name: 'blank' },
           { value: 344, name: 344 }].each do |occurrence|
            it "raises password configuration error for #{occurrence[:name]} password" do
              wrong_config = config
              wrong_config[:password] = occurrence[:value]

              expect do
                described_class.new(wrong_config)
              end.to raise_error(Sharepoint::Errors::InvalidNTLMConfigError)
            end
          end
        end
      end
    end
  end

  describe '#ethon_requester' do
    subject(:requester) { client.send(:ethon_easy_requester) }

    let(:client) { described_class.new(client_config) }
    let(:token) { instance_double(Token, access_token: 'footoken') }

    before do
      mock_token_responses
      allow(client).to receive(:authenticating_with_token).and_call_original
    end

    context 'when has token authentication' do
      let(:client_config) { sp_config(authentication: 'token') }

      it 'calls authenticating_with_token' do
        requester
        expect(client).to have_received(:authenticating_with_token)
      end

      it 'client token is set' do
        requester
        expect(client.token.access_token).not_to be_nil
      end
    end

    context 'when has ntlm authentication' do
      subject { client.send(:ethon_easy_requester) }

      let(:client_config) { sp_config(authentication: 'ntlm') }

      it 'does not call authenticating_with_token' do
        requester
        expect(client).not_to have_received(:authenticating_with_token)
      end

      it 'token is null' do
        expect(client.token.access_token).to be_nil
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
