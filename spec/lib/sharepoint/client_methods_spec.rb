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

  describe '#list_modified_documents' do
    before { mock_responses('list_modified_documents.json') }
    let(:list_name) { 'Documents' }
    let(:time) { Time.parse('2016-07-22') }
    subject { client.list_modified_documents time, list_name }
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
      it 'return documents modified after specified time' do
        results.each do |document|
          expect(Time.parse(document.modified)).to be >= time
        end
      end
    end
  end

  describe '#search_modified_documents' do
    let(:time) { Time.parse('2016-07-22') }
    let(:default_properties) do
      %w( write is_document list_id web_id created title author size path )
    end

    context 'search whole SharePoint instance' do
      before { mock_responses('search_modified_documents.json') }
      subject { client.search_modified_documents time }
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
        it 'return documents modified after specified time' do
          results.each do |document|
            expect(Time.parse(document.write)).to be >= time
          end
        end
        it 'return default properties with values' do
          sample = results.sample
          default_properties.each do |property|
            expect(sample).to respond_to property
            expect(sample.send(property)).not_to be_nil
          end
        end
      end
    end

    context 'search specific Site' do
      subject { client.search_modified_documents(time, options)[:results] }
      context 'when existing web_id is passed' do
        before { mock_responses('search_modified_documents.json') }
        let(:options) do
          { web_id: 'b285c5ff-9256-4f30-99ba-26fc705a9f2d' }
        end
        it { is_expected.not_to be_empty }
      end
      context 'when non-existing web_id is passed' do
        before { mock_responses('search_noresults.json') }
        let(:options) do
          { web_id: 'a285c5ff-9256-4f30-99ba-26fc705a9f2e' }
        end
        it { is_expected.to be_empty }
      end
    end

    context 'search specific List' do
      subject { client.search_modified_documents(time, options)[:results] }
      context 'when existing list_id is passed' do
        before { mock_responses('search_modified_documents.json') }
        let(:options) do
          { list_id: '3314c0cf-d5b0-4d1e-a5f1-9a10fca08bc3' }
        end
        it { is_expected.not_to be_empty }
      end
      context 'when non-existing list_id is passed' do
        before { mock_responses('search_noresults.json') }
        let(:options) do
          { list_id: '2314c0cf-d5b0-4d1e-a5f1-9a10fca08bc4' }
        end
        it { is_expected.to be_empty }
      end
    end

  end

  describe '#download' do
    let(:file_path) { '/Documents/document.docx' }
    let(:expected_content) { File.open('spec/fixtures/responses/document.docx').read }
    before { mock_responses('document.docx') }
    subject { client.download file_path }
    it { is_expected.to eq expected_content }
  end

  describe '#upload' do
    described_class::FILENAME_INVALID_CHARS.each do |char|
      it "shoud raise invalid file name error if the filename contains the character " + char do
        expect {
          client.upload(char + "filename", "content", "path")
        }.to raise_error(Sharepoint::Errors::InvalidSharepointFilename)
      end
    end
    # TODO
    it "should upload the file correctly"
  end

  # TODO
  describe ".update_metadata" do
    it "shoud raise invalid metadata if any metadata value or key include the single quote char"
    it "should update the metadata correctly"
  end

end
