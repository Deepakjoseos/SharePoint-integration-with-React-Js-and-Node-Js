// api.js
import axios from 'axios';

// Create a custom Axios instance with a base URL
const instance = axios.create({
  baseURL: `${process.env.REACT_APP_API_URL}`, // Set your backend base URL here
});

export default instance;
