// src/components/ProcessingPipeline.tsx

import React, { useState, useRef } from 'react';
import { FileText, Sparkles, X, AlertTriangle, CheckCircle, Folder, Download, File, Loader } from 'lucide-react';
import { saveAs } from 'file-saver';

interface ProcessResult {
  filename: string;
  result: {
    success: boolean;
    docx_filename?: string;
    docx_base64?: string;
    error?: string;
  };
  logs: string;
}

// Helper function to decode Base64 and create a Blob
const base64ToBlob = (base64: string, contentType: string) => {
  const byteCharacters = atob(base64);
  const byteArrays = [];
  for (let offset = 0; offset < byteCharacters.length; offset += 512) {
    const slice = byteCharacters.slice(offset, offset + 512);
    const byteNumbers = new Array(slice.length);
    for (let i = 0; i < slice.length; i++) {
      byteNumbers[i] = slice.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    byteArrays.push(byteArray);
  }
  return new Blob(byteArrays, { type: contentType });
};

const ProcessingPipeline = () => {
  const [files, setFiles] = useState<FileList | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [bulkResults, setBulkResults] = useState<ProcessResult[]>([]);
  const [selectedResult, setSelectedResult] = useState<ProcessResult | null>(null);
  const [activeTab, setActiveTab] = useState<'logs'>('logs');
  const [bulkMode, setBulkMode] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setFiles(e.target.files);
      setError(null);
      setBulkResults([]);
      setSelectedResult(null);
    }
  };

  const clearFiles = () => {
    setFiles(null);
    setBulkResults([]);
    setSelectedResult(null);
    setError(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const handleDownloadSingle = (item: ProcessResult) => {
    if (!item.result.success || !item.result.docx_base64) return;

    try {
      const blob = base64ToBlob(
        item.result.docx_base64,
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      );
      const outputFilename = item.result.docx_filename || `processed_${item.filename.replace(/\.md$/, '')}.docx`;
      saveAs(blob, outputFilename);
    } catch (e) {
      console.error("Failed to decode or save the DOCX file.", e);
      setError("Failed to download the DOCX file. It may be corrupted.");
    }
  };

  const handleDownloadAll = () => {
    bulkResults.forEach(item => {
      if (item.result.success) {
        handleDownloadSingle(item);
      }
    });
  };

  const handleProcess = async () => {
    if (!files || files.length === 0) {
      setError('Please upload a file or folder first.');
      return;
    }
    setIsLoading(true);
    setError(null);
    setBulkResults([]);
    setSelectedResult(null);

    const formData = new FormData();
    const isBulk = bulkMode || files.length > 1;

    if (isBulk) {
      Array.from(files).forEach(file => {
        formData.append('files', file);
      });
    } else {
      formData.append('file', files[0]);
    }

    const url = isBulk ? '/run_pipeline/bulk' : '/run_pipeline';

    try {
      const response = await fetch(url, { method: 'POST', body: formData });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || 'An unknown server error occurred.');
      }

      const results: ProcessResult[] = isBulk ? data.bulk_results : [data];

      setBulkResults(results);
      if (results.length > 0) {
        setSelectedResult(results[0]);
        setActiveTab('logs');
      }
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setIsLoading(false);
    }
  };

  const renderSelectedResult = () => {
    if (!selectedResult) return null;
    return (
      <div className="flex-1 lg:pl-6">
        <h3 className="text-lg font-semibold text-gray-900 mb-2 truncate">
          Details for: <span className="font-mono bg-gray-100 px-2 py-1 rounded">{selectedResult.filename}</span>
        </h3>
        <div className="border-b border-gray-200">
          <nav className="-mb-px flex space-x-6">
            <button
              onClick={() => setActiveTab('logs')}
              className={`whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm ${
                activeTab === 'logs'
                  ? 'border-blue-500 text-blue-600'
                  : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
              }`}
            >
              Processing Logs
            </button>
          </nav>
        </div>
        <div className="pt-4">
          {selectedResult.result.success ? (
            <>
              {activeTab === 'logs' && (
                <div>
                  <div className="p-4 bg-green-50 border border-green-200 rounded-lg mb-4">
                    <div className="flex items-center space-x-2">
                      <CheckCircle className="w-5 h-5 text-green-600" />
                      <span className="text-sm font-semibold text-green-800">
                        Conversion Successful
                      </span>
                    </div>
                    {selectedResult.result.docx_filename && (
                      <p className="text-sm text-green-700 mt-2">
                        Generated: {selectedResult.result.docx_filename}
                      </p>
                    )}
                  </div>
                  <div className="bg-gray-900 text-green-400 p-4 rounded-lg font-mono text-xs max-h-96 overflow-y-auto whitespace-pre-wrap">
                    {selectedResult.logs}
                  </div>
                </div>
              )}
            </>
          ) : (
            <div className="p-4 bg-red-50 text-red-800 border border-red-200 rounded-lg">
              <h4 className="font-semibold">Processing Failed</h4>
              <p className="text-sm mt-1">{selectedResult.result.error}</p>
              <h4 className="font-semibold mt-4">Logs</h4>
              <div className="bg-gray-900 text-red-400 p-2 mt-2 rounded font-mono text-xs max-h-64 overflow-y-auto whitespace-pre-wrap">
                {selectedResult.logs}
              </div>
            </div>
          )}
        </div>
      </div>
    );
  };

  const getFolderName = () => {
    if (!files || files.length === 0) return null;
    const firstFile = files[0] as any;
    const relativePath = firstFile.webkitRelativePath || '';
    return relativePath.split('/')[0] || `(${files.length} files)`;
  };

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
        {/* --- CONFIGURATION COLUMN --- */}
        <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6 space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="text-lg font-semibold text-gray-900">1. Upload & Configure</h3>
            <div className="flex items-center space-x-2">
              <label className="text-sm font-medium text-gray-700">Bulk Mode</label>
              <button
                onClick={() => setBulkMode(!bulkMode)}
                className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors ${
                  bulkMode ? 'bg-blue-600' : 'bg-gray-200'
                }`}
              >
                <span
                  className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform ${
                    bulkMode ? 'translate-x-6' : 'translate-x-1'
                  }`}
                />
              </button>
            </div>
          </div>

          {!files ? (
            <label
              htmlFor="md-upload"
              className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-lg cursor-pointer bg-gray-50 hover:bg-gray-100"
            >
              <div className="flex flex-col items-center justify-center">
                {bulkMode ? <Folder className="w-10 h-10 text-gray-400" /> : <File className="w-10 h-10 text-gray-400" />}
                <p className="text-sm text-gray-500 mt-2">
                  <span className="font-semibold">Click to select {bulkMode ? 'folder' : 'file'}</span> or drag & drop
                </p>
                <p className="text-xs text-gray-400 mt-1">Markdown files (.md) only</p>
              </div>
              <input
                id="md-upload"
                type="file"
                className="hidden"
                accept=".md,.markdown"
                onChange={handleFileChange}
                ref={fileInputRef}
                {...(bulkMode ? { webkitdirectory: 'true', mozdirectory: 'true' } : { multiple: true })}
              />
            </label>
          ) : (
            <div className="flex items-center justify-between p-3 bg-green-50 rounded-lg border border-green-200">
              <div className="flex items-center space-x-3">
                {bulkMode ? <Folder className="w-6 h-6 text-green-600" /> : <FileText className="w-6 h-6 text-green-600" />}
                <div>
                  <p className="text-sm font-medium text-gray-900">{bulkMode ? getFolderName() : files[0].name}</p>
                  <p className="text-xs text-green-700">{files.length} item(s) ready for processing.</p>
                </div>
              </div>
              <button onClick={clearFiles} className="p-1 hover:bg-gray-200 rounded-full transition-colors">
                <X className="w-4 h-4 text-gray-500" />
              </button>
            </div>
          )}

          <div className="p-4 bg-blue-50 border border-blue-200 rounded-lg">
            <h4 className="text-sm font-semibold text-blue-900 mb-2">Pipeline Steps:</h4>
            <ol className="text-xs text-blue-800 space-y-1 list-decimal list-inside">
              <li>Analyze markdown file status</li>
              <li>Convert MD files to DOCX format</li>
              <li>Add headers to DOCX documents</li>
              <li>Apply currency conversions</li>
              <li>Organize final output files</li>
            </ol>
          </div>
        </div>

        {/* --- ACTION & RESULTS COLUMN --- */}
        <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
          <h3 className="text-lg font-semibold text-gray-900 mb-4">2. Process & Review</h3>
          <button
            onClick={handleProcess}
            disabled={!files || isLoading}
            className="w-full flex items-center justify-center space-x-2 px-6 py-3 rounded-lg font-medium transition-all bg-blue-900 text-white hover:bg-blue-800 disabled:bg-gray-300 disabled:cursor-not-allowed"
          >
            {isLoading ? <Loader className="w-5 h-5 animate-spin" /> : <Sparkles className="w-5 h-5" />}
            <span>
              {isLoading ? `Converting ${files?.length} file(s)...` : `Convert ${files?.length || 0} file(s) to DOCX`}
            </span>
          </button>
          {error && (
            <div className="mt-4 p-4 bg-red-50 text-red-800 border border-red-200 rounded-lg flex items-start space-x-2">
              <AlertTriangle className="w-5 h-5 mt-0.5" />
              <div>
                <h4 className="font-semibold">Error</h4>
                <p className="text-sm">{error}</p>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* --- RESULTS PREVIEW AREA --- */}
      {bulkResults.length > 0 && (
        <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-lg font-semibold text-gray-900">Processing Results</h3>
            {bulkResults.some(r => r.result.success) && (
              <button
                onClick={handleDownloadAll}
                className="flex items-center space-x-2 px-4 py-2 text-sm font-medium rounded-lg bg-green-600 text-white hover:bg-green-700"
              >
                <Download className="w-4 h-4" />
                <span>Download All</span>
              </button>
            )}
          </div>
          <div className="flex flex-col lg:flex-row">
            <div className="w-full lg:w-1/3 lg:pr-6 border-b lg:border-b-0 lg:border-r border-gray-200 mb-4 lg:mb-0">
              <ul className="space-y-2 max-h-96 overflow-y-auto pr-2">
                {bulkResults.map((item) => (
                  <li key={item.filename}>
                    <div
                      className={`w-full text-left p-2 rounded-lg flex items-center justify-between transition-colors ${
                        selectedResult?.filename === item.filename ? 'bg-blue-100' : 'hover:bg-gray-100'
                      }`}
                    >
                      <button
                        onClick={() => setSelectedResult(item)}
                        className="flex items-center space-x-3 flex-grow truncate"
                      >
                        {item.result.success ? (
                          <CheckCircle className="w-5 h-5 text-green-500 flex-shrink-0" />
                        ) : (
                          <AlertTriangle className="w-5 h-5 text-red-500 flex-shrink-0" />
                        )}
                        <span className="text-sm font-medium text-gray-800 truncate">{item.filename}</span>
                      </button>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          handleDownloadSingle(item);
                        }}
                        disabled={!item.result.success}
                        className="p-2 rounded-md hover:bg-gray-200 disabled:text-gray-300 disabled:hover:bg-transparent"
                        aria-label={`Download ${item.filename}`}
                      >
                        <Download className="w-4 h-4 text-gray-600" title="Download as .docx" />
                      </button>
                    </div>
                  </li>
                ))}
              </ul>
            </div>
            {renderSelectedResult()}
          </div>
        </div>
      )}
    </div>
  );
};

export default ProcessingPipeline;