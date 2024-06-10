const axios = require('axios');
const fs = require('fs');
const formidable = require('formidable');
require('dotenv').config();
let DataverseaccessToken = '';
let SharePointaccessToken = null;
let tokenExpirationTime = 0; // Initialize token expiration time
let SPtokenExpirationTime = 0; // Initialize token expiration time

/*----------------------------------------------GET Sharepoint Access token Code Starts Here-------------------------------------------------------*/



const getSharePointAccessToken = async () => {
  // Check if the current access token is still valid
  if (SharePointaccessToken && Date.now() < SPtokenExpirationTime) {
    return SharePointaccessToken; // Return the existing token if it's still valid
  }

  try {
    // Retrieve environment variables from process.env
    const clientId = process.env.clientId;
    const clientSecret = process.env.client_secret;
    const tenantId = process.env.tenantId;
    const scope = process.env.SPscope;
    
    const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    // Send POST request to token endpoint to obtain access token
    const response = await axios.post(tokenEndpoint, new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: clientId,
      client_secret: clientSecret,
      scope: scope,
    }));

    // Validate the response
    if (!response.data.access_token || !response.data.expires_in) {
      throw new Error('Invalid response from token endpoint');
    }

    // Store the access token and its expiration time
    SharePointaccessToken = response.data.access_token;
    SPtokenExpirationTime = Date.now() + response.data.expires_in * 1000; // Convert seconds to milliseconds

    // Return the access token
    return SharePointaccessToken;
  } catch (error) {
    console.error('Error fetching access token:', error);
    throw new Error('Error fetching access token');
  }
};
/*----------------------------------------------GET Sharepoint Access token Code Ends Here-------------------------------------------------------*/


/*----------------------------------------------GET Dataverse Access token Code Starts Here-------------------------------------------------------*/
const getDataverseAccessToken = async () => {
  // Check if the current access token is still valid
  if (DataverseaccessToken && Date.now() < tokenExpirationTime) {
    return DataverseaccessToken; // Return the existing token if it's still valid
  }
  try {
    // Retrieve environment variables from process.env
    const clientId = process.env.clientId;
    const clientSecret = process.env.client_secret;
    const tenantId = process.env.tenantId;
    const resource = process.env.scope;

    const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    // Send POST request to token endpoint to obtain access token
    const response = await axios.post(
      tokenEndpoint,
      new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: clientId,
        client_secret: clientSecret,
        scope: resource,
      })
    );
    // Store the access token and its expiration time
    DataverseaccessToken = response.data.access_token;
    tokenExpirationTime = Date.now() + response.data.expires_in * 1000; // Convert seconds to milliseconds

    // Return the access token
    return DataverseaccessToken;
  } catch (error) {
    console.error('Error fetching access token:', error.response?.status, error.response?.data);
    throw new Error('Error fetching access token');
  }
};
/*----------------------------------------------GET Dataverse Access token Code Ends Here-------------------------------------------------------*/

/*----------------------------------------Get SharePoints Data from the SharePoint API Code Starts-----------------------------------------------*/
const getSharePointDocumentLocation = async (req, res) => {
  try {
    const { name } = req.params;

    // Ensure we have valid access tokens
    const [dataverseAccessToken, sharePointAccessToken] = await Promise.all([
      getDataverseAccessToken(),
      getSharePointAccessToken()
    ]);

    // Fetch SharePoint document locations
    const apiUrl = `${process.env.DataverseURL}/sharepointdocumentlocations?$filter=_regardingobjectid_value eq '${name}'`;
    const response = await axios.get(apiUrl, { headers: { Authorization: `Bearer ${dataverseAccessToken}` } });

    if (!response.data || !Array.isArray(response.data.value) || response.data.value.length === 0) {
      return res.status(404).json({ error: 'No document locations found' });
    }

    const firstDocumentLocation = response.data.value[0];
    if (!firstDocumentLocation || !firstDocumentLocation.relativeurl) {
      return res.status(404).json({ error: 'Relative URL not found in the first document location object' });
    }
    // Fetch document type choices and document details in parallel
    const [choiceResponse, documentsResponse] = await Promise.all([
      axios.get(`https://graph.microsoft.com/v1.0/sites/${process.env.SPSiteId}/lists/${process.env.SPListId}/columns`, { headers: { Authorization: `Bearer ${sharePointAccessToken}` } }),
      axios.get(`https://graph.microsoft.com/v1.0/drives/${process.env.SPDriveId}/root:/${firstDocumentLocation.relativeurl}:/children`, { headers: { Authorization: `Bearer ${sharePointAccessToken}` } })
    ]);

    if (!choiceResponse.data || !Array.isArray(choiceResponse.data.value)) {
      return res.status(500).json({ error: 'Error fetching document type choices' });
    }
    const folderId =documentsResponse.data.value[0].parentReference.id;
    const documentTypeColumn = choiceResponse.data.value.find(column => column.displayName === 'Document Type');
    if (!documentTypeColumn || !documentTypeColumn.choice) {
      return res.status(404).json({ error: 'Document Type column not found or it\'s not a choice column' });
    }

    const documentTypeChoices = documentTypeColumn.choice.choices;
    const documentType = documentTypeChoices;
    // console.log(documentType,"Types")

    if (!documentsResponse.data || !Array.isArray(documentsResponse.data.value)) {
      return res.status(500).json({ error: 'Error fetching document details' });
    }

    const documentsData = await Promise.all(documentsResponse.data.value.map(async item => {
      const metadataUrl = `https://graph.microsoft.com/v1.0/drives/${process.env.SPDriveId}/items/${item.id}/listItem`;
      const metadataResponse = await axios.get(metadataUrl, { headers: { Authorization: `Bearer ${sharePointAccessToken}` } });

      const customColumns = metadataResponse.data.fields || {};
      return {
        UniqueId: item.id,
        fileName: item.name,
        timeLastModified: item.lastModifiedDateTime,
        modifiedByTitle: item.lastModifiedBy ? item.lastModifiedBy.user.displayName : 'Unknown',
        downloadUrl: item["@microsoft.graph.downloadUrl"],
        documentType: customColumns.DocumentType,
        documentID: customColumns.DocumentID
      };
    }));
    
    // Extract fileDetails and documentType
    const fileDetails = documentsData.map(doc => ({
      UniqueId: doc.UniqueId,
      fileName: doc.fileName,
      timeLastModified: doc.timeLastModified,
      modifiedByTitle: doc.modifiedByTitle,
      downloadUrl: doc.downloadUrl,
      documentID: doc.documentID,
      documentType: doc.documentType,
    }));
    // console.log(fileDetails);
    // Send fileDetails and documentType to the frontend
    res.status(200).json({ fileDetails, documentType ,folderId });
  } catch (error) {
    console.error('Error fetching SharePoint requests:', error.response?.status, error.response?.data);
    res.status(500).json({ error: 'Internal server error' });
  }
};
/*--------------------------------------------Get SharePoints Data from the SharePoint API Code Ends----------------------------------------------*/



/*------------------------------------------------Upload File code Starts Here-------------------------------------------------------------------*/

const uploadToSharePoint = async (req, res) => {
  const form = new formidable.IncomingForm();
  form.parse(req, async (err, fields, files) => {
    if (err) {
      return res.status(400).send('Error parsing the files');
    }

    const file = files.file;
    if (!file) {
      return res.status(400).send('No file uploaded');
    }
    const filePath = file[0].filepath;
    const fileName = fields.filename[0];
    const folderPath = fields.FolderPath[0];
    const documentType = fields.documentType[0];
    const documentId = fields.documentID[0];
    
    try {
      const sharePointAccessToken = await getSharePointAccessToken();
      // Construct the upload URL
      const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${process.env.SPDriveId}/items/${folderPath}:/${fileName}:/content`;
      const fileBuffer = fs.readFileSync(filePath);
      // Set up the headers with the access token
      const headers = {
          'Authorization': `Bearer ${sharePointAccessToken}`,
          'Content-Type': 'application/json',
      };

      // Upload the file content using PUT request
      const response = await axios.put(uploadUrl, fileBuffer, { headers });
      // Construct the metadata update URL
      const metadataUrl = `https://graph.microsoft.com/v1.0/drives/${process.env.SPDriveId}/items/${folderPath}:/${fileName}:/listItem/fields`;
      
      // Payload for updating the document metadata
      const metadataPayload = {
        DocumentType: documentType,
        DocumentID: documentId
      };

      // Update the file metadata using PATCH request
      const metadataResponse = await axios.patch(metadataUrl, metadataPayload, { headers });
      
      res.status(200).send(`File '${fileName}' uploaded and metadata updated successfully`);    
    } catch (error) {
      console.error('Error uploading files to SharePoint:', error);
      res.status(500).send(`Failed to upload file '${fileName}' to SharePoint`);
    }
  });
};

/*----------------------------------------Upload File code Ends Here--------------------------------------------------------------------------*/

/*-------------------------------------Download SharePoint files Code Starts Here-------------------------------------------------------------*/
const downloadFromSharepoint = async (req, res) => { 
  try {
    // Extract id from query parameters
    const { id } = req.query;

    // Ensure we have a valid SharePoint access token
    const sharePointAccessToken = await getSharePointAccessToken();

    // Fetch the item metadata using the UniqueId
    const metadataUrl = `https://graph.microsoft.com/v1.0/drives/${process.env.SPDriveId}/items/${id}`;
    const metadataResponse = await axios.get(metadataUrl, { headers: { Authorization: `Bearer ${sharePointAccessToken}` } });

    if (!metadataResponse.data) {
      return res.status(404).json({ error: 'Document not found' });
    }

    // Get the download URL from the metadata response
    const downloadUrl = metadataResponse.data["@microsoft.graph.downloadUrl"];
    // Redirect to the download URL
    res.status(200).json({downloadUrl});
  } catch (error) {
    console.error('Error fetching SharePoint document:', error.response?.status, error.response?.data);
    res.status(500).json({ error: 'Internal server error' });
  }
};
/*-------------------------------------Download SharePoint files Code Ends Here-------------------------------------------------------------*/

const AddORUpdateDocumentType = async (req, res) => {
  const { fileName, newDocumentType, folderpath } = req.body;
  try {
    // Create an instance of SharePointFileUploader with your credentials
    const uploader = new SharePointFileUploader(process.env.SP_Username, process.env.SP_Password, process.env.SharePointURL);

    // Authenticate with SharePoint
    const authHeaders = await uploader.getAuthHeaders();
    const requestDigest = await uploader.getRequestDigest(authHeaders);

    // SharePoint API endpoint for updating a file
    const url = `${process.env.SharePointURL}/_api/Web/GetFolderByServerRelativeUrl('${folderpath}')/Files('${fileName}')/ListItemAllFields`;

    // Payload for the update
    const payload = {
      DocumentType: newDocumentType
    };

    // Make the POST request to update the item
    const response = await axios.post(url, payload, {
      headers: {
        ...authHeaders,
        'X-RequestDigest': requestDigest,
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json',
        'X-HTTP-Method': 'MERGE',
        'If-Match': '*'
      }
    });

    // Send response to client
    res.status(200).send('Update successful');
  } catch (error) {
    // Handle errors
    if (error.response) {
      // Server responded with a status other than 2xx
      console.error('Error response data:', error.response.data);
      console.error('Error response status:', error.response.status);
      console.error('Error response headers:', error.response.headers);
      res.status(error.response.status).send(`Error updating file: ${error.response.data.error.message}`);
    } else if (error.request) {
      // Request was made but no response was received
      console.error('Error request data:', error.request);
      res.status(500).send('Error updating file: No response from server');
    } else {
      // Something happened in setting up the request
      console.error('Error message:', error.message);
      res.status(500).send(`Error updating file: ${error.message}`);
    }
  }
};


/*----------------------------------------------Delete SharePoint File Code starts from here-------------------------------------------------------*/

const deleteFileFromSharePoint = async (req, res) => {
  try {
  const { id } = req.query;
  const { fileName } = req.query;
  console.log(id, "filepath");

  // Ensure we have a valid SharePoint access token
  const sharePointAccessToken = await getSharePointAccessToken();

  // Delete the item using the UniqueId
  const deleteUrl = `https://graph.microsoft.com/v1.0/drives/${process.env.SPDriveId}/items/${id}`;
  await axios.delete(deleteUrl, { headers: { Authorization: `Bearer ${sharePointAccessToken}` } });

  // Respond with success message

      res.status(200).send(`File ${fileName} deleted.`);
    } catch (error) {
      console.error('Error deleting SharePoint document:', error.response?.status, error.response?.data);
      res.status(500).json({ error: 'Internal server error' });
    }
  };


/*----------------------------------------------Delete SharePoint File Code starts ends here-------------------------------------------------------*/

module.exports = {
  getDataverseAccessToken,
  getSharePointDocumentLocation,
  uploadToSharePoint,
  downloadFromSharepoint,
  AddORUpdateDocumentType,
  deleteFileFromSharePoint,
  getSharePointAccessToken
};
