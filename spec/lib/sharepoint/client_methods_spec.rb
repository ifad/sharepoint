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
    subject { Sharepoint::Client.documents_for path }
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

end
