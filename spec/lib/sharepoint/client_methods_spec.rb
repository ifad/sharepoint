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
