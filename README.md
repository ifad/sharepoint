# sharepoint
Sharepoint 2013 REST API client. Work in progress, not for the faint hearted.

## Installation

Add this line to your application's Gemfile:

    gem 'sharepoint', git: 'git@github.com:ifad/sharepoint'

And then execute:

    bundle

## Usage

### Client initialization

You can instantiate a number of SharePoint clients in your application:

    client = Sharepoint::Client.new({
      username: "username",
      password: "password",
      uri: "http://sharepoint_url"
    })

### Get documents of a folder

    client.documents_for path

### Upload a document

    client.upload filename, content, path

### Update document metadata

    client.update_metadata filename, metadata, path
