import React, { useState } from 'react';
import axios from 'axios';

const FileUpload = () => {
    const [selectedFile, setSelectedFile] = useState(null);
    const [uploading, setUploading] = useState(false);
    const [uploadedData, setUploadedData] = useState(null);

    const handleFileChange = (event) => {
        setSelectedFile(event.target.files[0]);
    };

    const handleUpload = async () => {
        if (!selectedFile) return;

        setUploading(true);

        const formData = new FormData();
        formData.append("files", selectedFile);

        try {
            const response = await axios.post('http://localhost:8000/upload/', formData, {
                headers: {
                    'Content-Type': 'multipart/form-data'
                }
            });

            setUploadedData(response.data.files);
        } catch (error) {
            console.error('Error uploading file: ', error);
        } finally {
            setUploading(false);
        }
    };

    return (
        <div>
            <h2>Upload Excel File</h2>
            <input type="file" onChange={handleFileChange} />
            <button onClick={handleUpload} disabled={uploading}>Upload</button>

            {uploading && <p>Uploading...</p>}

            {uploadedData && (
                <div>
                    <h3>Uploaded Data</h3>
                    {uploadedData.map((file, index) => (
                        <div key={index}>
                            <h4>{file.filename}</h4>
                            <ul>
                                {file.data.map((row, idx) => (
                                    <li key={idx}>
                                        {row.column1} - {row.column2}
                                        {/* Display more columns as needed */}
                                    </li>
                                ))}
                            </ul>
                        </div>
                    ))}
                </div>
            )}
        </div>
    );
};

export default FileUpload;
