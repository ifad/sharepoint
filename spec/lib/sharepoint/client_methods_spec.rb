# frozen_string_literal: true

require 'spec_helper'

describe Sharepoint::Client do
  before do
    mock_requests
    mock_token_responses
  end

  let(:config) { sp_config }

  let(:client) { described_class.new(config) }

  describe '#documents_for' do
    subject(:documents_for) { client.documents_for path }

    let(:path) { '/Documents' }

    before { mock_responses('documents_for.json') }

    it 'returns documents with filled properties' do
      expect(documents_for).not_to be_empty
      sample = documents_for.sample
      %w[
        title path name url created_at updated_at
      ].each do |property|
        expect(sample).to respond_to property
        expect(sample.send(property)).not_to be_nil
      end
    end
  end

  describe '#get_document' do
    subject(:get_document) { client.get_document path }

    let(:path) { '/Documents/20160718 BRI-FCO boarding-pass.pdf' }

    before { mock_responses('get_document.json') }

    it { is_expected.to be_a OpenStruct }

    it 'returns expected document properties' do
      %w[guid title created modified].each do |property|
        expect(get_document).to respond_to property
        expect(get_document.send(property)).not_to be_nil
      end
    end

    it { is_expected.to respond_to(:url) }
  end

  describe '#document_exists?' do
    subject { client.document_exists? file_path, site_path }

    let(:file_path) { "#{list_path}#{folder_path}/#{file_name}" }
    let(:site_path) { '/sites/APRop' }

    context 'when list exists' do
      let(:list_path) { '/Lists/AFG' }

      context 'when folder exists' do
        let(:folder_path) { '/1100001460/Design Report' }

        context 'when file exists' do
          before { mock_responses('document_exists_true.json') }

          let(:file_name) { 'design_completion_part_1 without map.doc' }

          it { is_expected.to be true }
        end

        context "when file doesn't exist" do
          before { mock_responses('document_exists_false.json') }

          let(:file_name) { 'dummy.doc' }

          it { is_expected.to be false }
        end
      end

      context "when folder doesn't exist" do
        before { mock_responses('document_exists_false.json') }

        let(:folder_path) { '/foo/bar' }
        let(:file_name) { 'dummy.doc' }

        it { is_expected.to be false }
      end
    end

    context "when list doesn't exist" do
      before { mock_responses('document_exists_false.json') }

      let(:list_path) { '/Lists/foobar' }
      let(:folder_path) { '/1100001460/Design Report' }
      let(:file_name) { 'design_completion_part_1 without map.doc' }

      it { is_expected.to be false }
    end
  end

  describe '#list_documents' do
    subject(:list_documents) { client.list_documents list_name, conditions }

    before { mock_responses('list_documents.json') }

    let(:list_name) { 'Documents' }
    let(:time) { Time.parse('2016-07-22') }
    let(:conditions) { "Modified ge datetime'#{time.utc.iso8601}'" }

    it 'returns Hash with expected keys' do
      expect(list_documents).to be_a Hash
      expect(list_documents[:server_responded_at]).to be_a Time
      expect(list_documents[:results]).to be_a Array
    end

    describe 'results' do
      let(:results) { subject[:results] }

      it 'is not empty' do
        expect(results).not_to be_empty
      end

      it 'returns documents with filled properties' do
        sample = results.sample
        %w[
          unique_id title created modified name server_relative_url length
        ].each do |property|
          expect(sample).to respond_to property
          expect(sample.send(property)).not_to be_nil
        end
      end

      it 'returns documents verifying custom conditions' do
        expect(results.map { |document| Time.parse(document.modified) }).to all(be >= time)
      end

      it 'documents respond to url method' do
        expect(results).to all(respond_to :url)
      end
    end
  end

  describe '#search_modified_documents' do
    let(:start_at) { Time.parse('2016-07-24') }
    let(:end_at) { nil }
    let(:default_properties) do
      %w[write is_document list_id web_id created title author size path unique_id]
    end

    context 'when searching whole SharePoint instance' do
      subject(:search_modified_documents) { client.search_modified_documents({ start_at: start_at, end_at: end_at }) }

      before { mock_responses('search_modified_documents.json') }

      it 'returns Hash with expected keys' do
        expect(search_modified_documents).to be_a Hash
        expect(search_modified_documents[:server_responded_at]).to be_a Time
        expect(search_modified_documents[:results]).to be_a Array
      end

      describe 'results' do
        let(:results) { subject[:results] }

        it 'is not empty' do
          expect(results).not_to be_empty
        end

        it 'return document objects only' do
          expect(results.map(&:is_document)).to all(eq 'true')
        end

        it 'return documents modified after start_at' do
          expect(results.map { |document| Time.parse(document.write) }).to all(be >= start_at)
        end

        it 'return default properties with values' do
          sample = results.sample
          default_properties.each do |property|
            expect(sample).to respond_to property
            expect(sample.send(property)).not_to be_nil
          end
        end

        it 'documents respond to url method' do
          expect(results).to all(respond_to :url)
        end

        context 'with range end' do
          let(:end_at) { Time.parse('2016-07-26') }

          it 'return documents modified between start_at and end_at' do
            results.each do |document|
              modified_at = Time.parse(document.write)
              expect(modified_at).to be >= start_at
              expect(modified_at).to be <= end_at
            end
          end
        end
      end
    end

    context 'when searching specific Site' do
      subject { client.search_modified_documents(options)[:results] }

      let(:options) { { start_at: start_at } }

      context 'when existing web_id is passed' do
        before do
          mock_responses('search_modified_documents.json')
          options.merge!({ web_id: 'b285c5ff-9256-4f30-99ba-26fc705a9f2d' })
        end

        it { is_expected.not_to be_empty }
      end

      context 'when non-existing web_id is passed' do
        before do
          mock_responses('search_noresults.json')
          options.merge!({ web_id: 'a285c5ff-9256-4f30-99ba-26fc705a9f2e' })
        end

        it { is_expected.to be_empty }
      end
    end

    context 'when searching specific List' do
      subject { client.search_modified_documents(options)[:results] }

      let(:options) { { start_at: start_at } }

      context 'when existing list_id is passed' do
        before do
          mock_responses('search_modified_documents.json')
          options.merge!({ list_id: '3314c0cf-d5b0-4d1e-a5f1-9a10fca08bc3' })
        end

        it { is_expected.not_to be_empty }
      end

      context 'when non-existing list_id is passed' do
        before do
          mock_responses('search_noresults.json')
          options.merge!({ list_id: '2314c0cf-d5b0-4d1e-a5f1-9a10fca08bc4' })
        end

        it { is_expected.to be_empty }
      end
    end
  end

  describe '#search' do
    subject(:search) { client.search(options) }

    before { mock_responses('search_modified_documents.json') }

    let(:options) do
      {
        querytext: "'IsDocument=1'",
        refinementfilters: "'write:range(2016-07-23T22:00:00Z,max,from=\"ge\")'",
        selectproperties: "'Write,IsDocument,ListId,WebId,Created,Title,Author,Size,Path'",
        rowlimit: 500
      }
    end

    it { is_expected.not_to be_empty }

    context 'when request fails' do
      before do
        allow_any_instance_of(Ethon::Easy)
          .to receive(:response_code)
          .and_return(401)
      end

      it 'raises an error' do
        expect { search }.to raise_error(/\ARequest failed, received 401.+/)
      end
    end
  end

  describe '#download' do
    subject(:download) { client.download file_path: file_path }

    let(:file_path) { '/Documents/document.docx' }
    let(:expected_content) { File.read('spec/fixtures/responses/document.docx') }
    let(:document_meta) { client.send :parse_get_document_response, document_json, [] }

    context 'when meta contains URL' do
      let(:document_json) { File.read('spec/fixtures/responses/get_document.json') }

      before do
        allow(client).to receive(:get_document).and_return(document_meta)
        mock_responses('document.docx')
      end

      it 'returns expected hash' do
        expect(download).to have_key :file_contents
        expect(download).to have_key :link_url
        expect(download[:file_contents]).to eq expected_content
      end
    end

    context 'when meta contains Path' do
      let(:document_json) { File.read('spec/fixtures/responses/get_document_having_path.json') }

      before do
        allow(client).to receive(:get_document).and_return(document_meta)
        mock_responses('document.docx')
      end

      it 'returns expected hash' do
        expect(download).to have_key :file_contents
        expect(download).to have_key :link_url
        expect(download[:file_contents]).to eq expected_content
      end
    end
  end

  describe '#upload', pending: 'TODO' do
    it 'should upload the file correctly'
  end

  describe '.update_metadata', pending: 'TODO' do
    it 'shoud raise invalid metadata if any metadata value or key include the single quote char'
    it 'should update the metadata correctly'
  end

  describe '#folder_exists?' do
    specify do
      allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(200)
      expect(client.folder_exists?('foo')).to be(true)
    end

    specify do
      allow_any_instance_of(Ethon::Easy).to receive(:response_code).and_return(404)
      expect(client.folder_exists?('bar')).to be(false)
    end
  end

  describe '#create_folder' do
    it 'does nothing if the folder name is nil' do
      allow(Ethon::Easy).to receive(:new)
      expect(client.create_folder(nil, 'bar')).to be_nil
      expect(Ethon::Easy).not_to have_received(:new)
    end

    specify do
      mock_responses('request_digest.json')
      expect(client.create_folder('foo', 'bar')).to eq(200)
    end
  end

  describe '#lists' do
    subject(:lists) { client.lists(site_path, query) }

    before { mock_responses('lists.json') }

    let(:site_path) { '/sites/APRop' }
    let(:query) { { select: 'Title,Id,Hidden,ItemCount', filter: 'Hidden eq false' } }

    it 'returns Hash with expected keys' do
      expect(lists).to be_a Hash
      expect(lists[:server_responded_at]).to be_a Time
      expect(lists[:results]).to be_a Array
    end

    describe 'results' do
      let(:results) { subject[:results] }

      it 'is not empty' do
        expect(results).not_to be_empty
      end

      it 'returns lists with filled properties' do
        expect(lists).not_to be_empty
        sample = results.sample
        %w[
          hidden id item_count title
        ].each do |property|
          expect(sample).to respond_to property
          expect(sample.send(property)).not_to be_nil
        end
      end
    end
  end

  describe '#index_field' do
    subject(:index_field) { client.index_field('My List', 'Modified', site_path) }

    let(:site_path) { '/sites/APRop' }
    let(:response) { response_file }
    let(:response_file) { File.read('spec/fixtures/responses/index_field.json') }

    before do
      allow(client).to receive(:xrequest_digest).and_return('digest')

      allow_any_instance_of(Ethon::Easy)
        .to receive(:response_body)
        .and_return(response)

      allow_any_instance_of(Ethon::Easy)
        .to receive(:response_code)
        .and_return(204)
    end

    it 'updates the indexed field' do
      expect(index_field).to eq(204)
    end

    context 'when the field is already indexed' do
      let(:response) { JSON.parse(response_file).deep_merge('d' => { 'Indexed' => true }).to_json }

      it 'returns 304' do
        expect(index_field).to eq(304)
      end
    end
  end
end
