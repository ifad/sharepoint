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
  before { described_class.client = client }

  describe '#documents_for' do
    let(:path) { '/Documents' }
    before { mock_responses('documents_for.json') }
    subject { described_class.documents_for path }
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

  describe '#list_modified_documents' do
    before { mock_responses('list_modified_documents.json') }
    let(:list_name) { 'Documents' }
    let(:datetime) { Time.parse('2016-07-22') }
    subject { described_class.list_modified_documents datetime, list_name }
    it 'returns documents with filled properties' do
      is_expected.not_to be_empty
      sample = subject.sample
      %w(
        guid created modified name server_relative_url title
      ).each do |property|
        expect(sample).to respond_to property
        expect(sample.send(property)).not_to be_nil
      end
    end
    it 'returns documents modified after specified datetime' do
      subject.each do |document|
        expect(Time.parse(document.modified)).to be >= datetime
      end
    end
  end

  describe '#search_modified_documents' do
    let(:datetime) { Time.parse('2016-07-22') }
    let(:default_properties) do
      %w( write is_document list_id web_id title author size path )
    end

    context 'search whole SharePoint instance' do
      before { mock_responses('search_modified_documents.json') }
      subject { described_class.search_modified_documents datetime }
      it { is_expected.not_to be_empty }
      it 'returns document objects only' do
        subject.each do |document|
          expect(document.is_document).to eq 'true'
        end
      end
      it 'returns documents modified after specified datetime' do
        subject.each do |document|
          expect(Time.parse(document.write)).to be >= datetime
        end
      end
      it 'returns default properties' do
        sample = subject.sample
        default_properties.each do |property|
          expect(sample).to respond_to property
          expect(sample.send(property)).not_to be_nil
        end
      end
    end

    context 'search specific Site' do
      context 'when existing web_id is passed' do
        before { mock_responses('search_modified_documents.json') }
        let(:options) do
          { web_id: 'b285c5ff-9256-4f30-99ba-26fc705a9f2d' }
        end
        subject { described_class.search_modified_documents datetime, options }
        it { is_expected.not_to be_empty }
      end
      context 'when non-existing web_id is passed' do
        before { mock_responses('search_noresults.json') }
        let(:options) do
          { web_id: 'a285c5ff-9256-4f30-99ba-26fc705a9f2e' }
        end
        subject { described_class.search_modified_documents datetime, options }
        it { is_expected.to be_empty }
      end
    end

    context 'search specific List' do
      context 'when existing list_id is passed' do
        before { mock_responses('search_modified_documents.json') }
        let(:options) do
          { list_id: '3314c0cf-d5b0-4d1e-a5f1-9a10fca08bc3' }
        end
        subject { described_class.search_modified_documents datetime, options }
        it { is_expected.not_to be_empty }
      end
      context 'when non-existing list_id is passed' do
        before { mock_responses('search_noresults.json') }
        let(:options) do
          { list_id: '2314c0cf-d5b0-4d1e-a5f1-9a10fca08bc4' }
        end
        subject { described_class.search_modified_documents datetime, options }
        it { is_expected.to be_empty }
      end
    end

  end

end
