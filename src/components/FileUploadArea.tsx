import React, { useCallback, useState, useRef } from 'react'; // NEW: Added useState
import { Upload, Folder, X, File } from 'lucide-react';
import { useApp } from '../context/AppContext';

const FileUploadArea = () => {
  // Destructure all necessary values from the context
  const { setUploadedFiles, setSessionId, bulkMode, selectedMetric, selectedSubMetric, vendorAnalysisType, setVendorAnalysisType } = useApp();
  const [dragOver, setDragOver] = useState(false);
  const [folderName, setFolderName] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // A single, robust handler for uploading one or more files
  const handleFileUpload = useCallback(async (files: FileList | null) => {
    if (!files || files.length === 0) return;

    // Update UI to show what's selected
    if (files.length === 1) {
      setFileName(files[0].name);
    } else {
      setFileName(`${files.length} files selected`);
    }
    setFolderName(null); // Clear folder name if files are selected
    setUploadedFiles([]);
    setSessionId(null);

    // Create FormData and append all files
    const formData = new FormData();
    Array.from(files).forEach(file => {
      formData.append('files', file, file.name);
    });


    try {
      const response = await fetch('/upload', {
        method: 'POST',
        body: formData,
      });
      if (!response.ok) throw new Error('Server responded with an error!');
      const data = await response.json();
      setSessionId(data.session_id);
      console.log('Upload successful. Session ID:', data.session_id);
    } catch (error) {
      console.error('Error uploading file(s):', error);
      setFileName('Upload Failed!');
    }
  }, [setUploadedFiles, setSessionId]); // NEW: Added vendorAnalysisType to dependency array

  // Specific handler for folder selection, which then uses the main upload logic
  const onFolderSelect = useCallback(async (files: FileList | null) => {
    if (!files || files.length === 0) return;
    const firstFile = files[0];
    const relativePath = firstFile.webkitRelativePath || '';
    const folder = relativePath.split('/')[0];
    setFolderName(folder);
    setFileName(null);
    await handleFileUpload(files); // Reuse the core upload logic
  }, [handleFileUpload]);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
  };

  // Corrected drop handler to support multi-file drops
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer.files) {
      bulkMode ? onFolderSelect(e.dataTransfer.files) : handleFileUpload(e.dataTransfer.files);
    }
  };

  const clearSelection = () => {
    setFolderName(null);
    setFileName(null);
    setUploadedFiles([]);
    setSessionId(null);
    // --- NEW: Reset radio buttons to default ---
    setVendorAnalysisType('mom');
    // --- END NEW ---
    if (fileInputRef.current) {
        fileInputRef.current.value = "";
    }
  };

  const hasSelection = folderName || fileName;

  // Updated logic to show dynamic, helpful instructions
  const getDescriptionText = () => {
    if (selectedMetric === 'pex') {
      if (selectedSubMetric === 'pex-vendor') {
        return bulkMode
          ? 'Upload a root folder containing vendor sub-folders (each with 2024 & 2025 data).'
          : 'Upload the two vendor Excel files for 2024 and 2025.';
      }
      return 'Upload your PEX data file (e.g., PEX_US01_...xlsx).';
    }
    return 'Only .xlsx files will be processed.';
  };
  
  // Determines if multi-file selection should be enabled for the input
  const allowMultipleFiles = !bulkMode && selectedMetric === 'pex' && selectedSubMetric === 'pex-vendor';
  const buttonText = bulkMode ? 'Choose Folder' : allowMultipleFiles ? 'Choose Files' : 'Choose File';

  // --- NEW: Helper var to show/hide radio buttons ---
  const isPexVendor = selectedMetric === 'pex' && selectedSubMetric === 'pex-vendor';
  // --- END NEW ---

  return (
    <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
      <div className="flex items-center justify-between mb-4">
        <h3 className="text-lg font-semibold text-gray-900">
          {bulkMode ? 'Folder Upload' : 'File Upload'}
        </h3>
        {bulkMode && (
          <span className="px-3 py-1 text-xs font-medium bg-blue-100 text-blue-800 rounded-full">
            Bulk Mode
          </span>
        )}
      </div>

      {/* --- NEW: Conditionally rendered radio buttons --- */}
      {isPexVendor && (
        <div className="mb-4">
          <h4 className="text-sm font-medium text-gray-900 mb-2">Analysis Type</h4>
          <div className="flex gap-4">
            <label className="flex items-center space-x-2 cursor-pointer">
              <input
                type="radio"
                name="analysisType"
                value="mom"
                checked={vendorAnalysisType === 'mom'}
                onChange={() => setVendorAnalysisType('mom')}
                className="form-radio text-blue-900 focus:ring-blue-800"
              />
              <span className="text-sm text-gray-700">📅 Month-over-Month (MOM)</span>
            </label>
            <label className="flex items-center space-x-2 cursor-pointer">
              <input
                type="radio"
                name="analysisType"
                value="qtd"
                checked={vendorAnalysisType === 'qtd'}
                onChange={() => setVendorAnalysisType('qtd')}
                className="form-radio text-blue-900 focus:ring-blue-800"
              />
              <span className="text-sm text-gray-700">📊 Quarter-to-Date (QTD)</span>
            </label>
          </div>
        </div>
      )}
      {/* --- END NEW --- */}

      {!hasSelection ? (
        <div
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          className={`border-2 border-dashed rounded-lg p-8 text-center transition-colors ${
            dragOver
              ? 'border-blue-400 bg-blue-50'
              : 'border-gray-300 hover:border-gray-400'
          }`}
        >
          {bulkMode ? (
            <Upload className="w-12 h-12 text-gray-400 mx-auto mb-4" />
          ) : (
            <File className="w-12 h-12 text-gray-400 mx-auto mb-4" />
          )}
          <p className="text-lg font-medium text-gray-900 mb-2">
            {bulkMode ? 'Drop a folder here or click to select' : 'Drop file(s) here or click to select'}
          </p>
          <p className="text-sm text-gray-500 mb-4">
            {getDescriptionText()}
          </p>
          <input
            type="file"
            onChange={(e) => bulkMode ? onFolderSelect(e.target.files) : handleFileUpload(e.target.files)}
            className="hidden"
            id="file-upload"
            {...(bulkMode ? { webkitdirectory: "true", mozdirectory: "true", directory: "true" } : { multiple: allowMultipleFiles })}
            ref={fileInputRef}
          />
          <label
            htmlFor="file-upload"
            className="inline-flex items-center px-4 py-2 bg-blue-900 text-white rounded-lg hover:bg-blue-800 transition-colors cursor-pointer"
          >
            {buttonText}
          </label>
        </div>
      ) : (
        <div className="mt-6">
          <h4 className="text-sm font-medium text-gray-900 mb-3">Selected Item</h4>
          <div className="flex items-center justify-between p-3 bg-green-50 rounded-lg border border-green-200">
            <div className="flex items-center space-x-3">
              {folderName ? (
                <Folder className="w-6 h-6 text-green-600" />
              ) : (
                <File className="w-6 h-6 text-green-600" />
              )}
              <div>
                <p className="text-sm font-medium text-gray-900">{folderName || fileName}</p>
                <p className="text-xs text-green-700">Ready for processing.</p>
              </div>
            </div>
            <button
              onClick={clearSelection}
              className="p-1 hover:bg-gray-200 rounded-full transition-colors"
            >
              <X className="w-4 h-4 text-gray-500" />
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default FileUploadArea;