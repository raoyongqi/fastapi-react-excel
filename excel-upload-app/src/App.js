import React, { useState } from 'react';
import axios from 'axios';
import './App.css';

function App() {
  const [files, setFiles] = useState([]);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);

  const handleFileChange = (e) => {
    setFiles(e.target.files);
  };

  const handleUpload = async () => {
    const formData = new FormData();
    for (const file of files) {
      formData.append('files', file);
    }

    try {
      const response = await axios.post('http://127.0.0.1:8000/upload/', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      });
      setResult(response.data);
      setError(null);
    } catch (error) {
      console.error('Failed to upload files:', error);
      setError('Failed to upload files');
      setResult(null);
    }
  };

  return (
    <div className="container">
      <h1 className="header">Excel Upload App</h1>
      <div className="form-group">
        <input type="file" multiple onChange={handleFileChange} />
        <button className="button" onClick={handleUpload}>Upload</button>
      </div>
      {error && <p className="error">{error}</p>}
      {result && (
        <div className="result">
          <h2>Upload Results</h2>
          {Object.keys(result).map((fileName) => (
            <div key={fileName}>
              <h3>{fileName}</h3>
              <pre>{JSON.stringify(result[fileName], null, 2)}</pre>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export default App;
