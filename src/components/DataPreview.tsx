// import React, { useState, useEffect, useCallback } from 'react';
// import { Eye, EyeOff, Download, RefreshCw, Loader } from 'lucide-react';
// import { useApp } from '../context/AppContext';
 
// // Define types for fetched data row
// interface DataRow {
//     [key: string]: any;
// }
 
// const DataPreview = () => {
//     // Get sessionId and selectedMetric from App context
//     const { sessionId, selectedMetric } = useApp();
 
//     const [showPreview, setShowPreview] = useState(true);
//     const [previewType, setPreviewType] = useState<'raw' | 'cleaned'>('raw');
//     const [currentData, setCurrentData] = useState<DataRow[]>([]);
//     const [columnHeaders, setColumnHeaders] = useState<string[]>([]);
//     const [loading, setLoading] = useState(false);
//     const [error, setError] = useState<string | null>(null);
//     const [filename, setFilename] = useState<string | null>(null);
//     const [totalRows, setTotalRows] = useState<number | null>(null);
//     const [sheetName, setSheetName] = useState<string | null>(null);
//     const [debugInfo, setDebugInfo] = useState<string>('');
 
//     // Function to fetch data from the backend
//     const fetchData = useCallback(async () => {
//         if (!sessionId) {
//             setCurrentData([]);
//             setColumnHeaders([]);
//             setFilename(null);
//             setTotalRows(null);
//             setSheetName(null);
//             setDebugInfo('No session ID available');
//             return;
//         }
 
//         setLoading(true);
//         setError(null);
       
//         try {
//             // Pass the processing type (selectedMetric) to the backend
//             const processingType = selectedMetric || 'oe'; // Default to 'oe' if not set
//             const apiUrl = `/preview/${sessionId}?type=${previewType}&processing_type=${processingType}`;
           
//             setDebugInfo(`Fetching from: ${apiUrl}`);
//             console.log('Fetching preview data from:', apiUrl);
//             console.log('Session ID:', sessionId);
//             console.log('Selected Metric:', selectedMetric);
//             console.log('Preview Type:', previewType);
 
//             const response = await fetch(apiUrl, {
//                 method: 'GET',
//                 headers: {
//                     'Content-Type': 'application/json',
//                 },
//             });
 
//             console.log('Response status:', response.status);
//             console.log('Response ok:', response.ok);
 
//             const data = await response.json();
//             console.log('Response data:', data);
 
//             if (response.ok) {
//                 if (data.data && data.data.length > 0) {
//                         // Use the header row from the file if available
//                         let headers = [];
//                         if (Array.isArray(data.headers) && data.headers.length > 0) {
//                             headers = data.headers;
//                         } else {
//                             headers = Object.keys(data.data[0]);
//                         }
//                         // Move 'Unnamed' columns to the front, keep others in original order
//                         const sortedHeaders = [
//                             ...headers.filter((h: string) => h.startsWith('Unnamed')),
//                             ...headers.filter((h: string) => !h.startsWith('Unnamed'))
//                         ];
//                         setColumnHeaders(sortedHeaders);
//                     setCurrentData(data.data);
//                     setDebugInfo(`Successfully loaded ${data.data.length} rows`);
//                 } else {
//                     setColumnHeaders([]);
//                     setCurrentData([]);
//                     setDebugInfo('No data returned from server');
//                 }
//                 setFilename(data.filename);
//                 setTotalRows(data.total_rows);
//                 setSheetName(data.sheet_name);
//             } else {
//                 // Backend error (e.g., no files found)
//                 const errorMsg = data.error || 'Failed to fetch preview data.';
//                 setError(errorMsg);
//                 setDebugInfo(`Server error: ${errorMsg}`);
//                 setCurrentData([]);
//                 setColumnHeaders([]);
//                 setSheetName(null);
//             }
//         } catch (e) {
//             const errorMsg = `Network error: ${(e as Error).message}`;
//             console.error('Fetch error:', e);
//             setError('Network error or server unavailable.');
//             setDebugInfo(errorMsg);
//             setCurrentData([]);
//             setColumnHeaders([]);
//             setSheetName(null);
//         } finally {
//             setLoading(false);
//         }
//     }, [sessionId, previewType, selectedMetric]); // Added selectedMetric to dependencies
 
//     // Effect hook to fetch data whenever sessionId, previewType, or selectedMetric changes
//     useEffect(() => {
//         fetchData();
//     }, [fetchData]);
 
//     const handleRefresh = () => {
//         fetchData();
//     };
 
//     // Test server connection
//     const testConnection = async () => {
//         try {
//             const response = await fetch('/upload', {
//                 method: 'POST',
//                 body: new FormData() // Empty form data just to test connection
//             });
//             console.log('Server connection test:', response.status);
//             setDebugInfo(`Server connection test: ${response.status === 400 ? 'Connected (400 expected for empty request)' : `Status: ${response.status}`}`);
//         } catch (e) {
//             console.error('Connection test failed:', e);
//             setDebugInfo(`Connection test failed: ${(e as Error).message}`);
//         }
//     };
 
//     return (
//         <div className="bg-white rounded-lg shadow-sm border border-gray-200 mb-6">
//             <div className="p-6 border-b border-gray-200">
//                 <div className="flex items-center justify-between">
//                     <h3 className="text-lg font-semibold text-gray-900">
//                         Data Preview {filename && `(${filename}`}
//                         {sheetName && ` - Sheet: ${sheetName}`}
//                         {filename && ')'}
//                     </h3>
//                     <div className="flex items-center space-x-3">
//                         <div className="flex bg-gray-100 rounded-lg p-1">
//                             <button
//                                 onClick={() => setPreviewType('raw')}
//                                 disabled={loading}
//                                 className={`px-3 py-1 text-sm font-medium rounded-md transition-colors ${
//                                     previewType === 'raw'
//                                         ? 'bg-white text-gray-900 shadow-sm'
//                                         : 'text-gray-600 hover:text-gray-900'
//                                 }`}
//                             >
//                                 Raw Data
//                             </button>
//                             <button
//                                 onClick={() => setPreviewType('cleaned')}
//                                 disabled={loading}
//                                 className={`px-3 py-1 text-sm font-medium rounded-md transition-colors ${
//                                     previewType === 'cleaned'
//                                         ? 'bg-white text-gray-900 shadow-sm'
//                                         : 'text-gray-600 hover:text-gray-900'
//                                 }`}
//                             >
//                                 Cleaned Data
//                             </button>
//                         </div>
//                         <button
//                             onClick={() => setShowPreview(!showPreview)}
//                             className="flex items-center space-x-2 px-3 py-2 text-sm font-medium text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors"
//                         >
//                             {showPreview ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
//                             <span>{showPreview ? 'Hide' : 'Show'} Preview</span>
//                         </button>
//                     </div>
//                 </div>
               
//                 {/* Debug Information */}
//                 <div className="mt-4 p-3 bg-gray-50 rounded-lg">
//                     <div className="text-xs text-gray-600">
//                         <div><strong>Session ID:</strong> {sessionId || 'Not set'}</div>
//                         <div><strong>Selected Metric:</strong> {selectedMetric || 'Not set'}</div>
//                         <div><strong>Preview Type:</strong> {previewType}</div>
//                         <div><strong>Debug:</strong> {debugInfo}</div>
//                     </div>
//                     <button
//                         onClick={testConnection}
//                         className="mt-2 px-3 py-1 text-xs bg-blue-100 text-blue-700 rounded hover:bg-blue-200"
//                     >
//                         Test Server Connection
//                     </button>
//                 </div>
//             </div>
 
//             {showPreview && (
//                 <div className="p-6">
//                     {loading && (
//                         <div className="flex items-center justify-center p-8 text-blue-600">
//                             <Loader className="w-6 h-6 animate-spin mr-2" />
//                             Loading preview...
//                         </div>
//                     )}
 
//                     {error && (
//                          <div className="p-4 mb-4 text-sm text-red-700 bg-red-100 rounded-lg" role="alert">
//                              <span className="font-medium">Error:</span> {error}
//                          </div>
//                     )}
 
//                     {!loading && !error && currentData.length > 0 && (
//                         <>
//                             <div className="flex items-center justify-between mb-4">
//                                 <div className="flex items-center space-x-4">
//                                     <span className="text-sm text-gray-600">
//                                         Showing {currentData.length} of {totalRows || '...'} rows
//                                     </span>
//                                     {selectedMetric && (
//                                         <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-blue-100 text-blue-800">
//                                             {selectedMetric.toUpperCase()}
//                                         </span>
//                                     )}
//                                     {sheetName && (
//                                         <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">
//                                             Sheet: {sheetName}
//                                         </span>
//                                     )}
//                                 </div>
//                                 <div className="flex items-center space-x-2">
//                                     <button
//                                         onClick={handleRefresh}
//                                         className="flex items-center space-x-2 px-3 py-2 text-sm font-medium text-gray-700 hover:bg-gray-50 rounded-lg transition-colors"
//                                     >
//                                         <RefreshCw className="w-4 h-4" />
//                                         <span>Refresh</span>
//                                     </button>
//                                     <a
//                                         href={`/download/zip/${sessionId}`}
//                                         className="flex items-center space-x-2 px-3 py-2 text-sm font-medium text-blue-700 bg-blue-50 hover:bg-blue-100 rounded-lg transition-colors"
//                                     >
//                                         <Download className="w-4 h-4" />
//                                         <span>Export All</span>
//                                     </a>
//                                 </div>
//                             </div>
 
//                             <div className="overflow-x-auto">
//                                 <table className="w-full border border-gray-200 rounded-lg">
//                                     <thead className="bg-gray-50">
//                                         <tr>
//                                             {columnHeaders.map(header => (
//                                                 <th key={header} className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
//                                                     {header.replace(/_/g, ' ')}
//                                                 </th>
//                                             ))}
//                                         </tr>
//                                     </thead>
//                                     <tbody className="bg-white divide-y divide-gray-200">
//                                         {currentData.map((row, index) => (
//                                             <tr key={index} className="hover:bg-gray-50">
//                                                 {columnHeaders.map(header => (
//                                                     <td key={header} className="px-4 py-3 text-sm text-gray-900">
//                                                         {String(row[header] === null || row[header] === undefined ? 'N/A' : row[header])}
//                                                     </td>
//                                                 ))}
//                                             </tr>
//                                         ))}
//                                     </tbody>
//                                 </table>
//                             </div>
//                         </>
//                     )}
                   
//                     {!loading && !error && currentData.length === 0 && sessionId && (
//                         <div className="p-8 text-center text-gray-500">
//                             No data available for preview. Upload a file or run cleaning first.
//                         </div>
//                     )}
 
//                 </div>
//             )}
//         </div>
//     );
// };
 
// export default DataPreview;