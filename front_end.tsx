import React, { useState } from 'react';
import axios from 'axios';
import './App.css';

function App() {
    const [code, setCode] = useState('');
    const [documentation, setDocumentation] = useState(null);

    const handleUpload = async () => {
        try {
            const response = await axios.post('http://localhost:5000/parse_vba', { code });
            setDocumentation(response.data.documentation);
        } catch (error) {
            console.error('Error parsing VBA code:', error);
        }
    };

    return (
        <div className="App">
            <h1>VBA Code Parser</h1>
            <textarea
                value={code}
                onChange={(e) => setCode(e.target.value)}
                rows="10"
                cols="50"
                placeholder="Paste your VBA code here..."
            />
            <button onClick={handleUpload}>Parse</button>
            {documentation && (
                <pre>{JSON.stringify(documentation, null, 4)}</pre>
            )}
        </div>
    );
}

export default App;
