import React, { useState } from 'react';
import TextField from '@mui/material/TextField';
import Button from '@mui/material/Button';
import SendIcon from '@mui/icons-material/Send';
import SampleTab from './SampleTab';
import Snackbar from '@mui/material/Snackbar';
import Alert from '@mui/material/Alert';
import api from './api';

const MyForms = () => {
  const [name, setName] = useState('');
  const [sPData,setSPData] = useState(null);
  const [showSnackbar, setShowSnackbar] = useState(false);
  const [snackbarMessage, setSnackbarMessage] = useState('');
  const [snackbarSeverity, setSnackbarSeverity] = useState('success');

  const handleCloseSnackbar = () => {
    setShowSnackbar(false);
  };

  const getfundingrequests = async () => {
    try {
      const apiUrl2 = `/sharepointdocumentlocation/${name}`; // Corrected endpoint URL
  
    const sharepointResponse = await api.get(apiUrl2);
      
      const sharepointData = sharepointResponse.data;
      console.log(sharepointData,"spdata")
      setSPData(sharepointData)
      setTimeout(() => {
        setSnackbarMessage('Data retrieved successfully.');
        setSnackbarSeverity('success');
        setShowSnackbar(true);
      }, 2000);
      
    } catch (error) {
      console.error('Error fetching funding request:', error);
      setSnackbarMessage('Error retrieving data.');
      setSnackbarSeverity('error');
      setShowSnackbar(true);
    }
  };

  const handleInputChange = (event) => {
    setName(event.target.value); // Update the 'name' state when input changes
  };

  return (
    <>
      <TextField
        id="standard-basic"
        required
        variant="standard"
        value={name}
        onChange={handleInputChange}
        label="Application Number"
        name="applicationno"
        style={{ width: '20rem' }}
      />
      <Button
        onClick={getfundingrequests}
        variant="contained"
        style={{ marginTop: '0.7rem', marginLeft: '1rem' }}
        endIcon={<SendIcon />}
      >
        Retrieve Details
      </Button>

      {sPData && (
        <SampleTab sPData ={sPData} getfundingrequests={getfundingrequests}/>
      )}
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
    </>
  );
};

export default MyForms;
