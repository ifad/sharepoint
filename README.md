# sharepoint

[![Build Status](https://github.com/ifad/sharepoint/actions/workflows/ruby.yml/badge.svg)](https://github.com/ifad/sharepoint/actions)

Sharepoint 2013 REST API client. Work in progress, not for the faint hearted.

## Installation

Add this line to your application's Gemfile:

```rb
gem 'sharepoint', git: 'https://github.com/ifad/sharepoint.git'
```

And then execute:

    bundle

## Usage

### Client initialization

You can instantiate a number of SharePoint clients in your application:

```rb
client = Sharepoint::Client.new({
  username: 'username',
  password: 'password',
  uri: 'https://sharepoint_url'
})
```

### Get documents of a folder

```rb
client.documents_for path
```

### Upload a document

```rb
client.upload filename, content, path
```

### Update document metadata

```rb
client.update_metadata filename, metadata, path
```
