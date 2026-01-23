import { useState } from 'react';
import './FileUploader.css';

export default function FileUploader({ label, onFileSelect, accept = ".xlsx,.xls" }) {
  const [fileName, setFileName] = useState('');
  const [status, setStatus] = useState(''); // 'success', 'error', or ''

  const handleFileChange = async (e) => {
    const file = e.target.files[0];
    if (file) {
      setFileName(file.name);
      try {
        await onFileSelect(file);
        setStatus('success');
      } catch (error) {
        setStatus('error');
        alert(`文件解析失败: ${error.message}`);
      }
    }
  };

  const handleDrop = async (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) {
      setFileName(file.name);
      try {
        await onFileSelect(file);
        setStatus('success');
      } catch (error) {
        setStatus('error');
        alert(`文件解析失败: ${error.message}`);
      }
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
  };

  return (
    <div className="file-uploader">
      <label className="uploader-label">{label}</label>
      <div 
        className={`upload-area ${status}`}
        onDrop={handleDrop}
        onDragOver={handleDragOver}
      >
        <input 
          type="file" 
          accept={accept}
          onChange={handleFileChange}
          className="file-input"
        />
        <div className="upload-content">
          {fileName ? (
            <div className="file-info">
              <span className="file-name">{fileName}</span>
              {status === 'success' && <span className="status-icon">✓</span>}
              {status === 'error' && <span className="status-icon error">✗</span>}
            </div>
          ) : (
            <div className="upload-prompt">
              <span className="upload-icon">📁</span>
              <p>点击或拖拽文件到此处</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
