const express = require('express');
const {  getSharePointDocumentLocation, uploadToSharePoint , getDataverseAccessToken ,downloadFromSharepoint,AddORUpdateDocumentType, deleteFileFromSharePoint ,getSharePointAccessToken } = require('./controllers');

const router = express.Router();

// Define routes
router.get('/api/sharepointdocumentlocation/:name', getSharePointDocumentLocation);
router.post('/api/uploadToSharePoint', uploadToSharePoint);
router.get('/api/downloadfromSharepoint', downloadFromSharepoint);
router.delete('/api/deleteFileFromSharePoint',deleteFileFromSharePoint);





module.exports = {
  setupRoutes: (app) => {
    app.use(router);
  },
};
