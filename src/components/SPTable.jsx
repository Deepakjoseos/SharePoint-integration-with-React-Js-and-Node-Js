import React, { useState } from 'react';
import { DataGrid } from '@mui/x-data-grid';
import IconButton from '@mui/material/IconButton';
import DownloadIcon from '@mui/icons-material/Download';
import DeleteIcon from '@mui/icons-material/Delete';
import PictureAsPdfIcon from '@mui/icons-material/PictureAsPdf';
import moment from 'moment';
import InputLabel from '@mui/material/InputLabel'
import Button from '@mui/material/Button';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import { styled } from '@mui/material/styles';
import Alert from '@mui/material/Alert';
import Snackbar from '@mui/material/Snackbar';
import Select from '@mui/material/Select'; 
import MenuItem from '@mui/material/MenuItem';
import Dialog from '@mui/material/Dialog';
import DialogActions from '@mui/material/DialogActions';
import DialogContent from '@mui/material/DialogContent';
import DialogContentText from '@mui/material/DialogContentText';
import DialogTitle from '@mui/material/DialogTitle';
import { useTheme } from '@mui/material/styles';
import api from './api';

const VisuallyHiddenInput = styled('input')({
  clip: 'rect(0 0 0 0)',
  clipPath: 'inset(50%)',
  height: 1,
  overflow: 'hidden',
  position: 'absolute',
  bottom: 0,
  left: 0,
  whiteSpace: 'nowrap',
  width: 1,
});

const SPTable = ({ sPData,getfundingrequests }) => {
  const [selectedFile, setSelectedFile] = useState(null);
  const [showSnackbar, setShowSnackbar] = useState(false);
  const [snackbarMessage, setSnackbarMessage] = useState('');
  const [snackbarSeverity, setSnackbarSeverity] = useState('success');
  const [openDialog, setOpenDialog] = useState(false);
  const [selectedDocumentType, setSelectedDocumentType] = useState('');
  const [openDeleteDialog, setOpenDeleteDialog] = useState(false);
  const [selectedRow, setSelectedRow] = useState(null);
  const [documentID, setDocumentID] = useState('');
  const theme = useTheme();

  const handleFileChange = (event) => {
    setSelectedFile(event.target.files[0]); // Store the selected file
  };
  const handleCloseSnackbar = () => {
    setShowSnackbar(false);
};

const handleDocumentTypeChange = (event) => {
  setSelectedDocumentType(event.target.value);
};
const handleDocumentIDChange = (e) => {
  setDocumentID(e.target.value);
};
const handleUploadDocument = () => {
  setOpenDialog(true); // Open the dialog when "Upload Document" button is clicked
};

const handleDialogClose = () => {
  setOpenDialog(false); // Close the dialog
  setSelectedFile(null); // Reset the file input
  setSelectedDocumentType(''); // Reset the document type selection
  setDocumentID(''); //Rest the Document Id
};

const handleConfirmUpload = async () => {
  setOpenDialog(false); // Close the dialog
  try {
    if (!selectedFile) {
      setSnackbarMessage('Please select a file.');
      setSnackbarSeverity('error');
      setShowSnackbar(true);
      return;
    }
    if (!selectedDocumentType) {
      setSnackbarMessage('Please select a document type.');
      setSnackbarSeverity('error');
      setShowSnackbar(true);
      return;
    }
    if (!documentID) {
      setSnackbarMessage('Please enter a document ID.');
      setSnackbarSeverity('error');
      setShowSnackbar(true);
      return;
    }

    const formData = new FormData();
    formData.append('file', selectedFile);
    formData.append('filename', selectedFile.name);
    formData.append('FolderPath', sPData.folderId);
    formData.append('documentType', selectedDocumentType); // Add document type to form data
    formData.append('documentID', documentID);

    const response = await api.post('/uploadToSharePoint', formData, {
      headers: {
        'Content-Type': 'multipart/form-data',
      },
    });

    console.log('File uploaded successfully:', response.data);
    setSnackbarMessage(`${response.data}`);
    setSnackbarSeverity('success');
    setShowSnackbar(true);
    // After successful upload, refresh SharePoint data
    await getfundingrequests();
  } catch (error) {
    console.error('Error uploading file:', error);
    setSnackbarMessage(`${error.message}`);
    setSnackbarSeverity('error');
    setShowSnackbar(true);
  } finally {
    setSelectedFile(null); // Reset the selected file
    setSelectedDocumentType(''); // Reset the selected document type
    setDocumentID('');//Rest the Document Id
  }
};

//Downloading sharePoint file code starts from here

  const handleDownloadPdf = async (row) => {
    if (row && row.UniqueId) {
      try {  
        const response = await api.get(
          `/downloadfromSharepoint`,
          { params: { id: row.UniqueId }}
        );
        const downloadUrl = response.data.downloadUrl; // Ensure this is how you are accessing the URL
        console.log(downloadUrl,"url")
      // Check if the URL is valid and not empty
      if (downloadUrl) {
        // Create a link element to trigger the download
        const link = document.createElement('a');
        link.href = downloadUrl;

        // Append the link to the document body and trigger the download
        document.body.appendChild(link);
        link.click();

        // Clean up after the download is initiated
        link.remove();
        window.URL.revokeObjectURL(downloadUrl);
        setSnackbarMessage('File downloaded successfully.');
        setSnackbarSeverity('success');
        setShowSnackbar(true);
      } else {
        throw new Error('Invalid download URL');
      }
    } catch (error) {
      console.error('Error downloading file:', error);
      setSnackbarMessage(error);
      setSnackbarSeverity('error');
      setShowSnackbar(true);
    }
    }
  };
  
//Downloading sharePoint file code ends here  
  


//Delete File code starts
const handleDeleteFile = (row) => {
  setSelectedRow(row);
  setOpenDeleteDialog(true);
};

const handleDeleteDialogClose = () => {
  setOpenDeleteDialog(false);
  setSelectedRow(null);
};
const handleConfirmDelete = async () => {
  setOpenDeleteDialog(false);
  if (selectedRow && selectedRow.UniqueId) {
    try {
      const response = await api.delete('/deleteFileFromSharePoint', {
        params: { id: selectedRow.UniqueId, fileName: selectedRow.fileName }
      });

      console.log(response.data, "Delete Response");
      setSnackbarMessage(`${response.data}`);
      setSnackbarSeverity('success');
      setShowSnackbar(true);
      await getfundingrequests();
    } catch (error) {
      console.error('Error deleting file:', error);
      setSnackbarMessage('Failed to delete file.');
      setSnackbarSeverity('error');
      setShowSnackbar(true);
    }
  }
  setSelectedRow(null);
};
//Delete File code ends
//File Columns Starts
  const columns = [
    { field: 'fileName', headerName: 'Name', width: 350 ,headerClassName: 'header-bold',
    renderHeader: (params) => (
      <strong>{params.colDef.headerName}</strong>
    ),
    renderCell: (params) => (
      <div style={{ display: 'flex', alignItems: 'center' }}>
        <PictureAsPdfIcon color="error" />
        <span style={{ marginLeft: '5px' }}>{params.value}</span>
      </div>
    )
  },
    {
      field: 'timeLastModified',
      headerName: 'Modified',
      headerClassName: 'header-bold',
      renderHeader: (params) => (
        <strong>{params.colDef.headerName}</strong>
      ),
      width: 140,
      renderCell: (params) => moment(params.value).fromNow(),
    },
    {
      field: 'modifiedByTitle',
      headerName: 'Modified By',
      width: 130,
      renderHeader: (params) => (
        <strong>{params.colDef.headerName}</strong>
      )
    },
    {field: 'documentType', headerName: 'Document Type', width: 150 , headerClassName: 'header-bold',renderHeader: (params) => (
      <strong>{params.colDef.headerName}</strong>
    ) },
    {field: 'documentID', headerName: 'Document ID', width: 130 , headerClassName: 'header-bold',renderHeader: (params) => (
      <strong>{params.colDef.headerName}</strong>
    ) },    
    {
      field: 'Download',
      headerName: 'Actions',
      width: 110,
      sortable: false,
      renderHeader: (params) => (
        <strong>{params.colDef.headerName}</strong>
      ),
      renderCell: (params) => (
        <>
        <IconButton aria-label="download" onClick={() => handleDownloadPdf(params.row)}>
          <DownloadIcon color="primary" />
        </IconButton>
        <IconButton aria-label="delete" onClick={() => handleDeleteFile(params.row)}>
        <DeleteIcon color="error" />
      </IconButton>
        </>
      ),
    },
  ];
//File Columns Ends
  
  const getRowId = (row) => {
    if (row && row.UniqueId) {
      return row.UniqueId;
    }
    console.error('Invalid row or missing UniqueId:', row);
    return null;
  };

  return (
    <div style={{  width: '100%', display: 'flex', flexDirection: 'column' }}>
      <div style={{ marginBottom: '1rem', display: 'flex', alignItems: 'center' }}>
      <Button
          variant="contained"
          onClick={handleUploadDocument}
          style={{ marginLeft: '1rem' }}
        >
          Upload Document
        </Button>
        
        <Snackbar
          open={showSnackbar}
          autoHideDuration={5000}
          onClose={handleCloseSnackbar}
        >
          <Alert
            onClose={handleCloseSnackbar}
            severity={snackbarSeverity}
            sx={{ width: '100%' }}
          >
            {snackbarMessage}
          </Alert>
        </Snackbar>
      </div>

      <DataGrid
        rows={sPData.fileDetails}
        columns={columns}
        getRowId={getRowId}
        initialState={{
          pagination: {
            paginationModel: { page: 0, pageSize: 5 },
          },
        }}
        pageSizeOptions={[5,10]}
        pagination
        checkboxSelection
        autoHeight
        sx={{
          '& .header-bold': {
            fontWeight: 'bold',
            '& .MuiDataGrid-cell': {
              padding: '8px', // Adjust the padding as needed
            },
          },
        }}
      />
      {/* Popup To Add File and document Type---Code Starts */}
<Dialog
  open={openDialog}
  onClose={handleDialogClose}
  aria-labelledby="alert-dialog-title"
  aria-describedby="alert-dialog-description"
  fullWidth
  maxWidth="sm" // Adjust the maximum width as needed
>
  <DialogTitle id="alert-dialog-title">Upload Document</DialogTitle>
  <DialogContent>
    {/* Content for file selection */}
    <div style={{ marginBottom: '20px', display: 'flex', alignItems: 'center' }}>
      <Button
        component="label"
        role={undefined}
        variant="contained"
        startIcon={<CloudUploadIcon />}
      >
        Select File
        <VisuallyHiddenInput type="file" onChange={handleFileChange} />
      </Button>
      {selectedFile && (
        <div style={{ marginLeft: '10px', color: theme.palette.primary.main }}>
          {selectedFile.name}
        </div>
      )}
    </div>

    {/* Content for document type selection */}
    <div style={{ marginBottom: '20px', display: 'flex', alignItems: 'center' }}>
      <InputLabel id="document-type-label" style={{ fontWeight: 'bold', marginRight: '2.2rem' }}>
        Document Type
      </InputLabel>
      <Select
        labelId="document-type-label"
        id="document-type-select"
        value={selectedDocumentType}
        onChange={handleDocumentTypeChange}
        fullWidth
        style={{ flex: 1, fontSize: '0.9rem', height: '36px', maxWidth: '207px' }}
      >
        <MenuItem value="">Select Document Type</MenuItem>
        {sPData.documentType.map((type) => (
          <MenuItem key={type} value={type}>
            {type}
          </MenuItem>
        ))}
      </Select>
    </div>
    <div style={{ marginBottom: '20px', display: 'flex', alignItems: 'center' }}>
          <InputLabel id="document-id-label" style={{ fontWeight: 'bold', marginRight: '2.2rem' }}>
            Document ID
          </InputLabel>
          <input
            type="number"
            id="document-id-input"
            value={documentID}
            onChange={handleDocumentIDChange}
            style={{ flex: 1, fontSize: '0.9rem', height: '36px', maxWidth: '200px', marginLeft: '1.2rem' }}
          />
        </div>
  </DialogContent>
  <div style={{ padding: '20px', textAlign: 'left' }}>
    {/* Integrated buttons below the dialog content */}
    <Button onClick={handleDialogClose} component="label"
        role={undefined}
        variant="contained">
      Cancel
    </Button>
    <Button
      onClick={handleConfirmUpload}
      component="label"
      style={{ marginLeft: '0.8rem' }}
        role={undefined}
        variant="contained"
      disabled={!selectedFile || !selectedDocumentType || !documentID}
    >
      Upload
    </Button>
  </div>
</Dialog>
      {/* Popup To Add File and document Type---Code Ends */}

      {/* Delete File From the SharePoint---Code Starts */}
<Dialog
        open={openDeleteDialog}
        onClose={handleDeleteDialogClose}
        aria-labelledby="alert-dialog-title"
        aria-describedby="alert-dialog-description"
      >
        <DialogTitle id="alert-dialog-title">{"Confirm Delete"}</DialogTitle>
        <DialogContent>
          <DialogContentText id="alert-dialog-description">
            Are you sure you want to delete the file "{selectedRow?.fileName}"?
          </DialogContentText>
        </DialogContent>
        <DialogActions>
          <Button onClick={handleDeleteDialogClose} color="primary">
            Cancel
          </Button>
          <Button onClick={handleConfirmDelete} color="primary" autoFocus>
            Confirm
          </Button>
        </DialogActions>
      </Dialog>
      {/* Delete File From the SharePoint---Code Ends */}
      </div>
  );
};

export default SPTable;
