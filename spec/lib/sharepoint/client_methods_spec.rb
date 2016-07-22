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

end
