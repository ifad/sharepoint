require 'spec_helper'

describe Sharepoint::Client do
  before { mock_requests }
  let(:config) do
    {
      username: ENV['SP_USERNAME'],
      password: ENV['SP_PASSWORD'],
      uri: ENV['SP_URL']
    }
  end
  let(:client) { described_class.new(config) }

  describe '#documents_for' do
    let(:path) { '/Documents' }
    before { mock_responses('documents_for.json') }
    subject { client.documents_for path }
    it 'returns documents with filled properties' do
      is_expected.not_to be_empty
      sample = subject.sample
      %w(
        title path name url created_at updated_at
      ).each do |property|
        expect(sample).to respond_to property
        expect(sample.send(property)).not_to be_nil
      end
    end
  end

  describe '#get_document' do
    let(:path) { '/Documents/20160718 BRI-FCO boarding-pass.pdf' }
    before { mock_responses('get_document.json') }
    subject { client.get_document path }
    it { is_expected.to be_a OpenStruct }
    it 'returns expected document properties' do
      %w(guid title created modified).each do |property|
        expect(subject).to respond_to property
        expect(subject.send(property)).not_to be_nil
      end
    end
  end

  describe '#document_exists' do
    let(:file_path) { "#{list_path}#{folder_path}/#{file_name}" }
    let(:site_path) { '/sites/APRop' }

    subject { client.document_exists? file_path, site_path }

    context "when list exists" do
      let(:list_path) { "/Lists/AFG" }

      context "and folder exists" do
        let(:folder_path) { "/1100001460/Design Report" }

        context "and file exists" do
          before { mock_responses('document_exists_true.json') }
          let(:file_name) { "design_completion_part_1 without map.doc" }
          it { is_expected.to eq true }
        end

        context "and file doesn't exist" do
          before { mock_responses('document_exists_false.json') }
          let(:file_name) { "dummy.doc" }
          it { is_expected.to eq false }
        end
      end

      context "and folder doesn't exist" do
        before { mock_responses('document_exists_false.json') }
        let(:folder_path) { "/foo/bar" }
        let(:file_name) { "dummy.doc" }
        it { is_expected.to eq false }
      end
    end

    context "when list doesn't exist" do
      before { mock_responses('document_exists_false.json') }
      let(:list_path) { "/Lists/foobar" }
      let(:folder_path) { "/1100001460/Design Report" }
      let(:file_name) { "design_completion_part_1 without map.doc" }
      it { is_expected.to eq false }
    end
  end


  describe '#list_documents' do
    before { mock_responses('list_modified_documents.json') }
    let(:list_name) { 'Documents' }
    let(:time) { Time.parse('2016-07-22') }
    let(:conditions) { "Modified ge datetime'#{time.utc.iso8601}'" }
    subject { client.list_documents list_name, conditions }
    it 'returns Hash with expected keys' do
      expect(subject).to be_a Hash
      expect(subject[:server_responded_at]).to be_a Time
      expect(subject[:results]).to be_a Array
    end
    context 'results' do
      let(:results) { subject[:results] }
      it 'is not empty' do
        expect(results).not_to be_empty
      end
      it 'return documents with filled properties' do
        sample = results.sample
        %w(
          guid created modified name server_relative_url title
        ).each do |property|
          expect(sample).to respond_to property
          expect(sample.send(property)).not_to be_nil
        end
      end
      it 'return documents verifying custom conditions' do
        results.each do |document|
          expect(Time.parse(document.modified)).to be >= time
        end
      end
    end
  end

  describe '#search_modified_documents' do
    let(:start_at) { Time.parse('2016-07-24') }
    let(:end_at) { nil }
    let(:default_properties) do
      %w( write is_document list_id web_id created title author size path unique_id )
    end

    context 'search whole SharePoint instance' do
      before { mock_responses('search_modified_documents.json') }
      subject { client.search_modified_documents( { start_at: start_at, end_at: end_at } ) }
      it 'returns Hash with expected keys' do
        expect(subject).to be_a Hash
        expect(subject[:server_responded_at]).to be_a Time
        expect(subject[:results]).to be_a Array
      end
      context 'results' do
        let(:results) { subject[:results] }
        it 'is not empty' do
          expect(results).not_to be_empty
        end
        it 'return document objects only' do
          results.each do |document|
            expect(document.is_document).to eq 'true'
          end
        end
        it 'return documents modified after start_at' do
          results.each do |document|
            expect(Time.parse(document.write)).to be >= start_at
          end
        end
        it 'return default properties with values' do
          sample = results.sample
          default_properties.each do |property|
            expect(sample).to respond_to property
            expect(sample.send(property)).not_to be_nil
          end
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

    context 'search specific Site' do
      let(:options) { { start_at: start_at } }
      subject { client.search_modified_documents(options)[:results] }
      context 'when existing web_id is passed' do
        before do
          mock_responses('search_modified_documents.json')
          options.merge!( { web_id: 'b285c5ff-9256-4f30-99ba-26fc705a9f2d' } )
        end
        it { is_expected.not_to be_empty }
      end
      context 'when non-existing web_id is passed' do
        before do
          mock_responses('search_noresults.json')
          options.merge!( { web_id: 'a285c5ff-9256-4f30-99ba-26fc705a9f2e' } )
        end
        it { is_expected.to be_empty }
      end
    end

    context 'search specific List' do
      let(:options) { { start_at: start_at } }
      subject { client.search_modified_documents(options)[:results] }
      context 'when existing list_id is passed' do
        before do
          mock_responses('search_modified_documents.json')
          options.merge!( { list_id: '3314c0cf-d5b0-4d1e-a5f1-9a10fca08bc3' } )
        end
        it { is_expected.not_to be_empty }
      end
      context 'when non-existing list_id is passed' do
        before do
          mock_responses('search_noresults.json')
          options.merge!( { list_id: '2314c0cf-d5b0-4d1e-a5f1-9a10fca08bc4' } )
        end
        it { is_expected.to be_empty }
      end
    end

  end


  describe '#search' do
    before { mock_responses('search_modified_documents.json') }

    let(:options) do
      {
        querytext: "'IsDocument=1'",
        refinementfilters: "'write:range(2016-07-23T22:00:00Z,max,from=\"ge\")'",
        selectproperties: "'Write,IsDocument,ListId,WebId,Created,Title,Author,Size,Path'",
        rowlimit: 500
      }
    end

    subject { client.search(options) }

    it { is_expected.not_to be_empty }
  end

  describe '#download' do
    let(:file_path) { '/Documents/document.docx' }
    let(:expected_content) { File.open('spec/fixtures/responses/document.docx').read }
    before { mock_responses('document.docx') }
    subject { client.download file_path }
    it { is_expected.to eq expected_content }
  end

  describe '#upload' do
    # TODO
    it "should upload the file correctly"
  end

  # TODO
  describe ".update_metadata" do
    it "shoud raise invalid metadata if any metadata value or key include the single quote char"
    it "should update the metadata correctly"
  end

end
