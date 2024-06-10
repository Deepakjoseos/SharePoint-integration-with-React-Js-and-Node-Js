# Node.js Backend for React and SharePoint Integration

This repository hosts the Node.js backend for a web application that integrates a React frontend with SharePoint and Dataverse. The project aims to provide a robust solution for managing and interacting with organizational data through a modern, responsive interface.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [API Endpoints](#api-endpoints)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## Introduction

This Node.js backend provides the server-side logic for a web application that uses React for the frontend and integrates with SharePoint as the primary data source. It handles API requests, data processing, authentication, and communication with SharePoint and Dataverse.

## Features

- **Fetch SharePoint Data:** Retrieve document locations,metadata and files from SharePoint.
- **Upload Files to SharePoint:** Seamlessly upload files to SharePoint document libraries and update metadata.
- **Download Files from SharePoint:** Allow users to download files stored in SharePoint directly from the application.
- **Update Document Metadata:** Update document types and other metadata in SharePoint.
- **Delete Files:** Remove files from SharePoint.

## Prerequisites

Before you begin, ensure you have met the following requirements:
- Node.js (v12.x or later) and npm installed on your machine.
- Access to a SharePoint site and the necessary credentials.
- A Dataverse instance and the necessary API credentials.
- A `.env` file with the required environment variables (see Configuration).

## Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/Deepakjoseos/React-and-Node-Integration-with-SharePoint-and-Dataverse.git
    cd React-and-Node-Integration-with-SharePoint-and-Dataverse/backend
    ```

2. Install the dependencies:
    ```bash
    npm install
    ```

## Configuration

Create a `.env` file in the root of your project and add your own following environment variables:

```plaintext
PORT=5000
tenantId="Enter Azure Tenant Id"
clientId="Enter Azure Client Id"
client_secret="Enter Azure client secret"
scope="Enter Scope"
DataverseURL="Enter Dataverse URL"
SharePointURL="Enter SharePoint URL"
SP_Username="Enter username of SP"
SP_Password="Enter password of SP"
```

## Usage

1. Start the Node.js server:
```bash
npm start
```
2. The server will run on http://localhost:5000 (or another port if configured).

## API Endpoints

Get SharePoint Document Locations
URL: /api/sharepoint/documents/:name
Method: GET
Description: Fetches document locations from SharePoint based on the provided name.

Upload File to SharePoint
URL: /api/sharepoint/upload
Method: POST
Description: Uploads a file to SharePoint and updates its metadata.

Download File from SharePoint
URL: /api/sharepoint/download
Method: GET
Description: Downloads a file from SharePoint based on the provided server relative URL.
Query Parameters: serverRelativeUrl: The server relative URL of the file to download.


Update Document Metadata
URL: /api/sharepoint/update
Method: POST
Description: Updates the metadata of a document in SharePoint.

Delete File from SharePoint
URL: /api/sharepoint/delete
Method: DELETE
Description: Deletes a file from SharePoint based on the provided server relative URL.
Query Parameters:
serverRelativeUrl: The server relative URL of the file to delete.
fileName: The name of the file to delete.

## Troubleshooting

If you encounter issues, consider the following steps:

Ensure your .env file contains the correct credentials and URLs.
Check the console for error messages and stack traces.
Verify that you have the necessary permissions to access and modify data in SharePoint and Dataverse.

## Contributing

Contributions are welcome! Please fork the repository and create a pull request to propose changes.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
