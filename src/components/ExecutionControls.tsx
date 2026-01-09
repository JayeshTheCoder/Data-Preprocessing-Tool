import React, { useState } from 'react';
import { Play, Download, Terminal, CheckCircle, AlertCircle, Clock, Loader2 } from 'lucide-react';
import { useApp } from '../context/AppContext';

const ExecutionControls = () => {
  const { 
    bulkMode, 
    sessionId, 
    selectedMetric, 
    selectedSubMetric, 
    selectedWcType,
    selectedRules, // Already imported
    vendorAnalysisType 
  } = useApp();

  const [processing, setProcessing] = useState(false);
  const [showLogs, setShowLogs] = useState(false);
  const [progress, setProgress] = useState(0);
  const [logs, setLogs] = useState<string[]>([]);
  const [cleanedFiles, setCleanedFiles] = useState<string[]>([]);

  const handleStartProcessing = async () => {
    if (!sessionId) {
      alert("Please upload a file or folder first.");
      return;
    }
    
    let endpointPath = '';
    let fetchBody: { [key: string]: any } = {};
    let logMetricName = selectedMetric.toUpperCase();

    // --- UPDATED PAYLOAD ---
    // This now spreads all rules (e.g., validateFormats) from the context
    // and merges them with other necessary flags.
    const basePayload = {
      ...selectedRules, // Spreads { removeDuplicates: bool, groupUnits: bool, ... }
      bulk_mode: bulkMode
    };
    // --- END UPDATE ---

    switch (selectedMetric) {
      case 'sales':
        endpointPath = '/clean_sales/';
        fetchBody = basePayload;
        break;
      case 'oe':
        endpointPath = '/clean_oe/';
        fetchBody = basePayload;
        break;
      case 'pex':
        endpointPath = '/clean_pex/';
        fetchBody = { 
          ...basePayload, 
          sub_metric: selectedSubMetric,
          vendorAnalysisType: vendorAnalysisType 
        };
        logMetricName = `PEX (${selectedSubMetric?.toUpperCase()})`;
        break;
      case 'working-capital':
        endpointPath = '/clean_wc/';
        fetchBody = { 
          ...basePayload, 
          metric: selectedWcType 
        };
        logMetricName = `WC (${selectedWcType.toUpperCase()})`;
        break;
      default:
        alert(`No processing logic defined for metric: ${selectedMetric}`);
        return;
    }
    
    const apiUrl = `${endpointPath}${sessionId}`;
    
    setProcessing(true);
    setProgress(0);
    setCleanedFiles([]);
    setLogs([`Starting data processing for ${logMetricName}...`]);

    const interval = setInterval(() => {
      setProgress(prev => Math.min(prev + 10, 90));
    }, 200);

    try {
      const fetchOptions: RequestInit = {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json' 
        },
        body: JSON.stringify(fetchBody) // Send the complete payload
      };

      const response = await fetch(apiUrl, fetchOptions);
      const data = await response.json();
      clearInterval(interval);

      if (response.ok) {
        setCleanedFiles(data.cleaned_files);
        setProgress(100);
        setLogs(prev => [...prev, `Processing complete. ${data.cleaned_files.length} file(s) created.`]);
      } else {
        throw new Error(data.error || `Processing failed.`);
      }
    } catch (error) {
      console.error('Error processing files:', error);
      setLogs(prev => [...prev, `Error: ${(error as Error).message}`]);
      clearInterval(interval);
    } finally {
      setProcessing(false);
    }
  };

  const handleDownload = (filename: string) => {
    window.location.href = `/download/${sessionId}/${filename}`;
  };

  const handleDownloadAll = () => {
    if (!sessionId) return;
    window.location.href = `/download/zip/${sessionId}`;
  };

  const processingStatus = processing
    ? { icon: Clock, color: 'text-yellow-600', bg: 'bg-yellow-50', border: 'border-yellow-200' }
    : progress === 100
    ? { icon: CheckCircle, color: 'text-green-600', bg: 'bg-green-50', border: 'border-green-200' }
    : { icon: AlertCircle, color: 'text-gray-400', bg: 'bg-gray-50', border: 'border-gray-200' };

  return (
    <div className="bg-white rounded-lg shadow-sm border border-gray-200">
      <div className="p-6 border-b border-gray-200">
        <h3 className="text-lg font-semibold text-gray-900 mb-4">
          Execution Controls
        </h3>

        <div className="flex items-center justify-between mb-6">
          <div className="flex space-x-3">
            <button
              onClick={handleStartProcessing}
              disabled={processing || !sessionId}
              className={`flex items-center space-x-2 px-6 py-3 rounded-lg font-medium transition-all ${
                processing || !sessionId
                  ? 'bg-gray-200 text-gray-500 cursor-not-allowed'
                  : 'bg-blue-900 text-white hover:bg-blue-800 active:transform active:scale-95'
              }`}
            >
              {processing ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                <Play className="w-4 h-4" />
              )}
              <span>{processing ? 'Processing...' : 'Run Process'}</span>
            </button>
          </div>

          <div className="flex items-center space-x-3">
            <button
              onClick={() => setShowLogs(!showLogs)}
              className="flex items-center space-x-2 px-4 py-2 text-sm font-medium text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors"
            >
              <Terminal className="w-4 h-4" />
              <span>{showLogs ? 'Hide' : 'Show'} Logs</span>
            </button>
          </div>
        </div>

        {/* Progress Section */}
        <div className={`p-4 rounded-lg ${processingStatus.bg} ${processingStatus.border} border`}>
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center space-x-2">
              <processingStatus.icon className={`w-5 h-5 ${processingStatus.color}`} />
              <span className={`font-medium ${processingStatus.color}`}>
                {processing ? 'Processing...' : progress === 100 ? 'Complete' : 'Ready to Process'}
              </span>
            </div>
            <span className="text-sm text-gray-600">{progress}%</span>
          </div>

          <div className="w-full bg-gray-200 rounded-full h-2">
            <div
              className={`h-2 rounded-full transition-all duration-300 ${
                processing ? 'bg-yellow-500' : progress === 100 ? 'bg-green-500' : 'bg-gray-300'
              }`}
              style={{ width: `${progress}%` }}
            />
          </div>

          {progress === 100 && (
            <div className="mt-3 text-sm text-green-700">
              Processing completed successfully. Files ready for download.
            </div>
          )}
        </div>
      </div>

      {/* Download Section */}
      {cleanedFiles.length > 0 && (
        <div className="p-6 border-t border-gray-200">
          
          <div className="flex items-center justify-between mb-4">
            <h4 className="text-lg font-semibold text-gray-900">
              Download Processed Files
            </h4>
            <button
              onClick={handleDownloadAll}
              className="flex-shrink-0 flex items-center space-x-2 px-4 py-2 text-sm font-medium rounded-lg bg-green-600 text-white hover:bg-green-700"
            >
              <Download className="w-4 h-4" />
              <span>Download All (.zip)</span>
            </button>
          </div>

          <div className="space-y-2">
            {cleanedFiles.map((filename) => (
              <div key={filename} className="flex items-center justify-between p-3 bg-gray-50 rounded-lg">
                <p className="text-sm font-medium text-gray-900 truncate pr-4">{filename}</p>
                <button
                  onClick={() => handleDownload(filename)}
                  className="flex-shrink-0 flex items-center space-x-2 px-4 py-2 text-sm font-medium rounded-lg bg-green-600 text-white hover:bg-green-700"
                >
                  <Download className="w-4 h-4" />
                  <span>Download</span>
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Console Logs */}
      {showLogs && (
        <div className="p-6 border-t border-gray-200">
          <h4 className="text-sm font-medium text-gray-900 mb-3">Processing Logs</h4>
          <div className="bg-gray-900 text-green-400 p-4 rounded-lg font-mono text-sm max-h-64 overflow-y-auto">
            {logs.map((log, index) => (
              <div key={index} className="mb-1">
                {log}
              </div>
            ))}
            {processing && (
              <div className="flex items-center space-x-2">
                <div className="w-2 h-2 bg-green-400 rounded-full animate-pulse" />
                <span>Processing in progress...</span>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default ExecutionControls;