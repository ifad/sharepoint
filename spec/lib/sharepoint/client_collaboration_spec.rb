# frozen_string_literal: true

require 'spec_helper'

RSpec.describe Sharepoint::Client, 'Collaboration Methods' do
  before do
    ENV['SP_URL'] = 'https://localhost:8888'
    ENV['SP_AUTHENTICATION'] = 'token'
    ENV['SP_CLIENT_ID'] = 'test_client_id'
    ENV['SP_CLIENT_SECRET'] = 'test_client_secret'
    ENV['SP_TENANT_ID'] = 'test_tenant_id'
    ENV['SP_CERT_NAME'] = 'test_cert'
    ENV['SP_AUTH_SCOPE'] = 'https://test.sharepoint.com/.default'

    mock_requests
    mock_token_responses

    # Mock xrequest_digest calls for POST requests
    allow_any_instance_of(described_class).to receive(:xrequest_digest).and_return('mock_digest_value')
  end

  let(:config) do
    sp_config(authentication: 'token').merge({
                                               site_path: '/sites/test-site',
                                               base_folder: 'Shared Documents/Collaboration',
                                               base_uri: 'https://company.sharepoint.com'
                                             })
  end

  let(:client) { described_class.new(config) }
  let(:filename) { 'test-document.docx' }
  let(:folder_path) { 'project-123' }
  let(:file_content) { 'test file content' }

  describe '#collaboration_upload_document' do
    subject(:upload_document) { client.collaboration_upload_document(filename, file_content, folder_path) }

    let(:unique_id) { '12345678-1234-1234-1234-123456789abc' }
    let(:upload_response) { { 'd' => { 'UniqueId' => unique_id } }.to_json }

    before do
      allow(client).to receive(:update_metadata)

      # Mock collaboration_upload_and_get_unique_id to return the unique_id
      allow(client).to receive_messages(folder_exists?: true, collaboration_upload_and_get_unique_id: unique_id)
    end

    it 'returns success with unique_id' do
      expect(upload_document).to eq({ success: true, unique_id: unique_id })
    end

    context 'with metadata' do
      subject(:upload_document) do
        client.collaboration_upload_document(filename, file_content, folder_path, { 'Title' => 'Test Document' })
      end

      it 'updates metadata' do
        upload_document
        expect(client).to have_received(:update_metadata).with(
          filename,
          { 'Title' => 'Test Document' },
          'Shared Documents/Collaboration/project-123',
          '/sites/test-site'
        )
      end
    end

    context 'when folder does not exist' do
      before do
        allow(client).to receive(:folder_exists?).and_return(false)
        allow(client).to receive(:create_folder)
      end

      it 'creates the folder' do
        upload_document
        expect(client).to have_received(:create_folder).with(
          folder_path,
          'Shared Documents/Collaboration',
          '/sites/test-site'
        )
      end
    end

    context 'when upload fails' do
      before do
        allow(client).to receive(:collaboration_upload_and_get_unique_id).and_raise(StandardError, 'Upload failed')
      end

      it 'raises UploadError' do
        expect { upload_document }.to raise_error(Sharepoint::Errors::UploadError, /Failed to upload document/)
      end
    end
  end

  describe '#collaboration_grant_permissions' do
    subject(:grant_permissions) do
      client.collaboration_grant_permissions(filename, folder_path, user_emails, role_id)
    end

    let(:user_emails) { ['user1@example.com', 'user2@example.com'] }
    let(:role_id) { 1_073_741_827 } # Contribute role
    let(:user_id_1) { 123 }
    let(:user_id_2) { 456 }

    before do
      allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(200)
      allow_any_instance_of(Ethon::Easy).to receive(:response_body).and_return(
        { 'd' => { 'Id' => user_id_1 } }.to_json
      )
    end

    it 'returns true on success' do
      expect(grant_permissions).to be true
    end

    context 'with nil user_emails' do
      let(:user_emails) { nil }

      it 'raises PermissionError' do
        expect { grant_permissions }.to raise_error(
          Sharepoint::Errors::PermissionError,
          'user_emails cannot be nil'
        )
      end
    end

    context 'with non-array user_emails' do
      let(:user_emails) { 'user@example.com' }

      it 'raises PermissionError' do
        expect { grant_permissions }.to raise_error(
          Sharepoint::Errors::PermissionError,
          'user_emails must be an array'
        )
      end
    end

    context 'with empty user_emails' do
      let(:user_emails) { [] }

      it 'raises PermissionError' do
        expect { grant_permissions }.to raise_error(
          Sharepoint::Errors::PermissionError,
          'user_emails cannot be empty'
        )
      end
    end

    context 'with nil role_id' do
      let(:role_id) { nil }

      it 'raises PermissionError' do
        expect { grant_permissions }.to raise_error(
          Sharepoint::Errors::PermissionError,
          'role_id is required'
        )
      end
    end
  end

  describe '#collaboration_revoke_user_permission' do
    subject(:revoke_user_permission) do
      client.collaboration_revoke_user_permission(filename, folder_path, user_email)
    end

    let(:user_email) { 'user@example.com' }
    let(:user_id) { 123 }

    before do
      allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(200)
      allow_any_instance_of(Ethon::Easy).to receive(:response_body).and_return(
        { 'd' => { 'Id' => user_id } }.to_json
      )
    end

    it 'returns true on success' do
      expect(revoke_user_permission).to be true
    end

    context 'when revocation fails' do
      before do
        allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(500)
      end

      it 'raises PermissionError' do
        expect { revoke_user_permission }.to raise_error(Sharepoint::Errors::PermissionError)
      end
    end
  end

  describe '#collaboration_revoke_all_permissions' do
    subject(:revoke_all_permissions) do
      client.collaboration_revoke_all_permissions(filename, folder_path)
    end

    before do
      allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(200)
    end

    it 'returns true on success' do
      expect(revoke_all_permissions).to be true
    end

    context 'when reset fails' do
      before do
        allow(client).to receive(:collaboration_reset_role_inheritance).and_raise(StandardError, 'Reset failed')
      end

      it 'raises PermissionError' do
        expect { revoke_all_permissions }.to raise_error(
          Sharepoint::Errors::PermissionError,
          /Failed to revoke all permissions/
        )
      end
    end
  end

  describe '#collaboration_get_web_edit_url' do
    subject(:get_web_edit_url) do
      client.collaboration_get_web_edit_url(filename, folder_path, unique_id)
    end

    let(:unique_id) { '12345678-1234-1234-1234-123456789abc' }

    context 'with unique_id provided' do
      it 'returns a web edit URL with GUID format' do
        expect(get_web_edit_url).to include('/_layouts/15/Doc.aspx')
        expect(get_web_edit_url).to include(unique_id.upcase)
        expect(get_web_edit_url).to include('test-document.docx')
      end
    end

    context 'without unique_id' do
      let(:unique_id) { nil }

      before do
        allow(client).to receive(:collaboration_get_file_unique_id).and_return(nil)
      end

      it 'returns a fallback folder URL' do
        expect(get_web_edit_url).to include('Shared')
        expect(get_web_edit_url).to include('Documents')
        expect(get_web_edit_url).to include('Collaboration')
        expect(get_web_edit_url).to include('project-123')
      end
    end

    context 'with empty unique_id' do
      let(:unique_id) { '' }

      it 'returns a fallback folder URL' do
        expect(get_web_edit_url).to include('Shared')
        expect(get_web_edit_url).to include('Documents')
        expect(get_web_edit_url).to include('Collaboration')
        expect(get_web_edit_url).to include('project-123')
      end
    end
  end

  describe '#detect_office_type' do
    {
      'document.doc' => 'w',
      'document.docx' => 'w',
      'spreadsheet.xls' => 'x',
      'spreadsheet.xlsx' => 'x',
      'presentation.ppt' => 'p',
      'presentation.pptx' => 'p',
      'other.pdf' => 'd'
    }.each do |filename, expected_type|
      it "detects #{expected_type} for #{filename}" do
        expect(client.detect_office_type(filename)).to eq(expected_type)
      end
    end
  end

  describe '#collaboration_get_file_unique_id' do
    subject(:get_file_unique_id) { client.collaboration_get_file_unique_id(server_relative_url) }

    let(:server_relative_url) { '/sites/test-site/Shared Documents/file.docx' }
    let(:unique_id) { '12345678-1234-1234-1234-123456789abc' }

    before do
      allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(200)
      allow_any_instance_of(Ethon::Easy).to receive(:response_body).and_return(
        { 'd' => { 'UniqueId' => unique_id } }.to_json
      )
    end

    it 'returns the unique_id' do
      expect(get_file_unique_id).to eq(unique_id)
    end

    context 'when file not found' do
      before do
        allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(404)
      end

      it 'returns nil' do
        expect(get_file_unique_id).to be_nil
      end
    end

    context 'when API call fails' do
      before do
        allow_any_instance_of(Ethon::Easy).to receive(:perform).and_raise(StandardError, 'Network error')
      end

      it 'returns nil' do
        expect(get_file_unique_id).to be_nil
      end
    end
  end

  describe '#collaboration_download_document' do
    subject(:download_document) { client.collaboration_download_document(filename, folder_path) }

    let(:file_contents) { 'downloaded file content' }

    before do
      allow(client).to receive(:download).and_return({ file_contents: file_contents })
    end

    it 'returns file contents' do
      expect(download_document).to eq(file_contents)
    end

    context 'when download method returns string directly' do
      before do
        allow(client).to receive(:download).and_return(file_contents)
      end

      it 'returns the string' do
        expect(download_document).to eq(file_contents)
      end
    end

    context 'when download fails' do
      before do
        allow(client).to receive(:download).and_raise(StandardError, 'Download failed')
      end

      it 'raises DownloadError' do
        expect { download_document }.to raise_error(
          Sharepoint::Errors::DownloadError,
          /Failed to download document/
        )
      end
    end
  end

  describe '#collaboration_delete_document' do
    subject(:delete_document) { client.collaboration_delete_document(filename, folder_path) }

    before do
      allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(200)
    end

    it 'returns true on success' do
      expect(delete_document).to be true
    end

    context 'when deletion fails' do
      before do
        allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(404)
      end

      it 'raises DeleteError' do
        expect { delete_document }.to raise_error(
          Sharepoint::Errors::DeleteError,
          /Failed to delete document: HTTP 404/
        )
      end
    end

    context 'when API call raises error' do
      before do
        allow_any_instance_of(Ethon::Easy).to receive(:perform).and_raise(StandardError, 'Network error')
      end

      it 'raises DeleteError' do
        expect { delete_document }.to raise_error(
          Sharepoint::Errors::DeleteError,
          /Failed to delete document/
        )
      end
    end
  end

  describe '#collaboration_delete_folder' do
    subject(:delete_folder) { client.collaboration_delete_folder(folder_path) }

    before do
      allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(200)
    end

    it 'returns true on success' do
      expect(delete_folder).to be true
    end

    context 'when deletion fails' do
      before do
        allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(404)
        allow_any_instance_of(Ethon::Easy).to receive(:response_body).and_return('Not Found')
      end

      it 'raises DeleteError' do
        expect { delete_folder }.to raise_error(
          Sharepoint::Errors::DeleteError,
          /Failed to delete folder: HTTP 404/
        )
      end
    end

    context 'when API call raises error' do
      before do
        allow_any_instance_of(Ethon::Easy).to receive(:perform).and_raise(StandardError, 'Network error')
      end

      it 'raises DeleteError' do
        expect { delete_folder }.to raise_error(
          Sharepoint::Errors::DeleteError,
          /Failed to delete folder/
        )
      end
    end
  end

  describe '#collaboration_document_exists?' do
    subject(:collaboration_document_exists?) do
      client.collaboration_document_exists?(filename, folder_path)
    end

    before do
      allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(200)
    end

    it 'returns true when document exists' do
      expect(collaboration_document_exists?).to be true
    end

    context 'when document does not exist' do
      before do
        allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(404)
      end

      it 'returns false' do
        expect(collaboration_document_exists?).to be false
      end
    end

    context 'when API call raises error' do
      before do
        allow_any_instance_of(Ethon::Easy).to receive(:perform).and_raise(StandardError, 'Network error')
      end

      it 'returns false' do
        expect(collaboration_document_exists?).to be false
      end
    end
  end

  describe 'private helper methods' do
    describe '#collaboration_site_path' do
      it 'returns site_path from config' do
        expect(client.send(:collaboration_site_path)).to eq('/sites/test-site')
      end

      context 'when site_path is not configured' do
        let(:config) { sp_config(authentication: 'token').merge({ base_folder: 'Shared Documents' }) }

        it 'returns empty string' do
          expect(client.send(:collaboration_site_path)).to eq('')
        end
      end
    end

    describe '#collaboration_base_folder' do
      it 'returns base_folder from config' do
        expect(client.send(:collaboration_base_folder)).to eq('Shared Documents/Collaboration')
      end

      context 'when base_folder is not configured' do
        let(:config) { sp_config(authentication: 'token').merge({ site_path: '/sites/test' }) }

        it 'returns empty string' do
          expect(client.send(:collaboration_base_folder)).to eq('')
        end
      end
    end

    describe '#collaboration_base_uri' do
      it 'returns base_uri from config' do
        expect(client.send(:collaboration_base_uri)).to eq('https://company.sharepoint.com')
      end

      context 'when base_uri is not configured' do
        let(:config) do
          sp_config(authentication: 'token').merge({
                                                     uri: 'https://default-uri.sharepoint.com',
                                                     site_path: '/sites/test',
                                                     base_folder: 'Shared'
                                                   })
        end

        it 'falls back to uri from config' do
          expect(client.send(:collaboration_base_uri)).to eq('https://default-uri.sharepoint.com')
        end
      end
    end

    describe '#collaboration_build_full_path' do
      it 'combines base_folder and folder_path' do
        full_path = client.send(:collaboration_build_full_path, 'project-123')
        expect(full_path).to eq('Shared Documents/Collaboration/project-123')
      end

      context 'when base_folder is empty' do
        let(:config) { sp_config(authentication: 'token').merge({ site_path: '/sites/test' }) }

        it 'returns only folder_path' do
          full_path = client.send(:collaboration_build_full_path, 'project-123')
          expect(full_path).to eq('project-123')
        end
      end

      context 'when folder_path is empty' do
        it 'returns only base_folder' do
          full_path = client.send(:collaboration_build_full_path, '')
          expect(full_path).to eq('Shared Documents/Collaboration')
        end
      end
    end

    describe '#collaboration_build_server_relative_url' do
      it 'builds complete server-relative URL' do
        url = client.send(:collaboration_build_server_relative_url, filename, folder_path)
        expect(url).to eq('/sites/test-site/Shared Documents/Collaboration/project-123/test-document.docx')
      end
    end
  end

  describe 'configuration validation' do
    context 'with missing site_path' do
      let(:config) do
        sp_config(authentication: 'token').merge({
                                                   base_folder: 'Shared Documents',
                                                   base_uri: 'https://company.sharepoint.com'
                                                 })
      end

      it 'uses empty string as default' do
        expect(client.send(:collaboration_site_path)).to eq('')
      end
    end

    context 'with missing base_folder' do
      let(:config) do
        sp_config(authentication: 'token').merge({
                                                   site_path: '/sites/test-site',
                                                   base_uri: 'https://company.sharepoint.com'
                                                 })
      end

      it 'uses empty string as default' do
        expect(client.send(:collaboration_base_folder)).to eq('')
      end
    end

    context 'with missing base_uri' do
      let(:config) do
        sp_config(authentication: 'token').merge({
                                                   site_path: '/sites/test-site',
                                                   base_folder: 'Shared Documents'
                                                 })
      end

      it 'falls back to uri' do
        expect(client.send(:collaboration_base_uri)).to eq(ENV.fetch('SP_URL', nil))
      end
    end
  end
end
