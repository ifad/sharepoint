# sharepoint
Sharepoint 2013 REST API client. Work in progress, not for the faint hearted.

## Installation

Add this line to your application's Gemfile:

    gem 'sharepoint', git: 'git@github.com:ifad/sharepoint'

And then execute:

    bundle

## Usage

  # client configuration

    Sharepoint::Client.client = Sharepoint::Client.new({username: "username", password: "password", uri: "http://sharepoint_url"})

  # get documents of a folder

    Sharepoint::Client.documents_for path

  # upload document

    Sharepoint::Client.upload filename, content, path

  # update document metadata

    Sharepoint::Client.update_metadata filename, metadata, path
