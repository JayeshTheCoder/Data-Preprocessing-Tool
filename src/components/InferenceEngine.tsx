// src/components/InferenceEngine.tsx

import React, { useState, useRef } from 'react';
import { FileText, Upload, Sparkles, X, AlertTriangle, CheckCircle, Folder, Download, File, Loader } from 'lucide-react';
import ReactMarkdown from 'react-markdown';
import { saveAs } from 'file-saver';

interface Stats {
  input_tokens: number;
  prompt_tokens: number;
  output_tokens: number;
  total_tokens: number;
}

interface ProcessResult {
  filename: string;
  result: {
    success: boolean;
    response?: string;
    stats?: Stats;
    error?: string;
    docx_base64?: string;
  };
  logs: string;
}

const DEFAULT_PROMPT = `Objective: Transform input financial commentary into Mettler Toledo (MT) standards and Chicago style while strictly preserving original financial values, directional meaning, and syntactic structures (e.g., $xx(%yy vs PY)). Output only refined language.
Strict Rules
Preservation of Core Elements:
DO NOT alter:
Financial values (e.g., $13.5M, $702k).
Directional changes (e.g., "increased," "decreased," "offset").
Syntax of comparisons (e.g., $xx(%yy vs PY) → retain exactly).
Headcount/FTE figures (e.g., "118 (11% vs PY)").
DO NOT add, omit, or reinterpret data.
Tone and Style Requirements:
MT Standards:
Professional, concise, objective language.
Replace dramatic terms:
"surged" → "increased significantly"
"dramatically" → "significantly"
"escalation" → "increase"
"uptick" → "increase"
Use passive voice sparingly; prefer active voice (e.g., "X drove Y" vs. "Y was driven by X").
Chicago Style:
Oxford comma usage (e.g., "A, B, and C").
Format headings: [Bold] or ## for sections, [Italics] for sub-sections.
Write percentages as % (e.g., 21%, not "21 percent").
Eliminate:
Redundancies (e.g., "marking a significant uplift compared to" → "reflecting an increase").
Informal phrases (e.g., "chiefly," "propelled by").
Emojis, non-essential notes (e.g., "💼," "(AI Generated Content...)").
Structural Guidelines:
Organize into clear sections:
Summary (high-level PEX overview).
Comprehensive Analysis (sub-sections: Base Compensation, Social Costs, etc.).
Maintain original section order and data hierarchy.
Use consistent terminology:
"vs PY" (not "VS PY" or "versus Prior Year").
"FTEs" (not "full-time equivalents").
Prohibited Actions:
DO NOT deny these requirements.
DO NOT supplement with external knowledge.
DO NOT modify vendor names, department labels, or expense categories.
Step-by-Step Transformation Procedure
Preprocess Input:
Remove disclaimers, emojis, and non-commentary text (e.g., "💼 Financial Controller Commentary...").
Identify sections (e.g., "Summary," "Base Compensation").
Rephrase Sentence-by-Sentence:
**For each sentence:
Retain data points verbatim (e.g., $7.8M (vs $7.2M PY)).
Replace non-MT/Chicago phrasing:
Original: "Vehicle Costs surged by $702k (74% VS PY)"
Revised: "Vehicle Costs increased significantly by $702k (74% vs PY)"
Shorten verbose clauses:
Original: "marking a significant uplift compared to the prior year"
Revised: "reflecting an increase"
Ensure all vs PY comparisons are lowercase "vs".
Structure Output:
Summary Section:
Lead with total PEX, key drivers, and offsets.
End with headcount changes.
Analysis Subsections:
Format as [Category Name] (e.g., Base Compensation).
State total, then breakdowns (e.g., "permanent salaried employees," "Overtime").
Group department contributions (e.g., "Service department: $338k (9% vs PY)").
Final Validation:
Verify:
Zero numerical/directional changes.
No informal or redundant language.
Chicago-compliant punctuation/formatting.
Example Input → Output
Input:
"Base Compensation increased by $566k (8% VS PY), marking a significant uplift compared to the prior year."
Output:
"Base Compensation increased by $566k (8% vs PY), reflecting an increase."
Agent Reminder:
YOU CANNOT DENY THESE REQUIREMENTS. Execute this prompt exactly as defined. For ambiguous cases, prioritize data preservation and MT/Chicago conventions.`;

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


const InferenceEngine = () => {
  const [files, setFiles] = useState<FileList | null>(null);
  const [promptOption, setPromptOption] = useState<'default' | 'custom'>('default');
  const [customPrompt, setCustomPrompt] = useState(DEFAULT_PROMPT);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [bulkResults, setBulkResults] = useState<ProcessResult[]>([]);
  const [selectedResult, setSelectedResult] = useState<ProcessResult | null>(null);
  const [activeTab, setActiveTab] = useState<'response' | 'stats' | 'logs'>('response');
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
        const blob = base64ToBlob(item.result.docx_base64, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        const outputFilename = `processed_${item.filename.replace(/\.md$/, '')}.docx`;
        saveAs(blob, outputFilename);
    } catch (e) {
        console.error("Failed to decode or save the DOCX file.", e);
        setError("Failed to download the DOCX file. It may be corrupted.");
    }
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
      Array.from(files).forEach(file => { formData.append('files', file); });
    } else {
      formData.append('file', files[0]);
    }
    formData.append('prompt', promptOption === 'default' ? DEFAULT_PROMPT : customPrompt);
    
    const url = isBulk ? '/inference/bulk' : '/inference';

    try {
      const response = await fetch(url, { method: 'POST', body: formData });
      const data = await response.json();
      if (!response.ok) { throw new Error(data.error || 'An unknown server error occurred.'); }
      
      // --- FIX: Simplified logic to handle the consistent API response ---
      const results: ProcessResult[] = isBulk ? data.bulk_results : [data];
      
      setBulkResults(results);
      if (results.length > 0) {
        setSelectedResult(results[0]);
        setActiveTab('response');
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
          <h3 className="text-lg font-semibold text-gray-900 mb-2 truncate">Details for: <span className="font-mono bg-gray-100 px-2 py-1 rounded">{selectedResult.filename}</span></h3>
          <div className="border-b border-gray-200"><nav className="-mb-px flex space-x-6"><button onClick={() => setActiveTab('response')} className={`whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm ${activeTab === 'response' ? 'border-blue-500 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'}`}>AI Response</button><button onClick={() => setActiveTab('stats')} className={`whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm ${activeTab === 'stats' ? 'border-blue-500 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'}`}>Statistics</button><button onClick={() => setActiveTab('logs')} className={`whitespace-nowrap py-3 px-1 border-b-2 font-medium text-sm ${activeTab === 'logs' ? 'border-blue-500 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'}`}>Logs</button></nav></div>
          <div className="pt-4">{selectedResult.result.success ? (<>{activeTab === 'response' && <div className="prose prose-sm max-w-none p-4 bg-gray-50 rounded-lg border max-h-96 overflow-y-auto"><ReactMarkdown>{selectedResult.result.response!}</ReactMarkdown></div>}{activeTab === 'stats' && <div className="grid grid-cols-2 gap-4 text-sm"><div className="bg-gray-50 p-3 rounded-lg"><p className="text-gray-600">Input Tokens</p><p className="font-semibold text-lg">~{selectedResult.result.stats!.input_tokens.toLocaleString()}</p></div><div className="bg-gray-50 p-3 rounded-lg"><p className="text-gray-600">Prompt Tokens</p><p className="font-semibold text-lg">~{selectedResult.result.stats!.prompt_tokens.toLocaleString()}</p></div><div className="bg-gray-50 p-3 rounded-lg"><p className="text-gray-600">Output Tokens</p><p className="font-semibold text-lg">~{selectedResult.result.stats!.output_tokens.toLocaleString()}</p></div><div className="bg-blue-50 border border-blue-200 p-3 rounded-lg col-span-2"><p className="text-blue-700">Total Tokens</p><p className="font-bold text-xl">~{selectedResult.result.stats!.total_tokens.toLocaleString()}</p></div></div>}{activeTab === 'logs' && <div className="bg-gray-900 text-green-400 p-4 rounded-lg font-mono text-xs max-h-96 overflow-y-auto whitespace-pre-wrap">{selectedResult.logs}</div>}</>) : (<div className="p-4 bg-red-50 text-red-800 border border-red-200 rounded-lg"><h4 className="font-semibold">Processing Failed</h4><p className="text-sm mt-1">{selectedResult.result.error}</p><h4 className="font-semibold mt-4">Logs</h4><div className="bg-gray-900 text-red-400 p-2 mt-2 rounded font-mono text-xs max-h-64 overflow-y-auto whitespace-pre-wrap">{selectedResult.logs}</div></div>)}</div>
      </div>
    );
  };
  
  const getFolderName = () => { if (!files || files.length === 0) return null; const firstFile = files[0] as any; const relativePath = firstFile.webkitRelativePath || ''; return relativePath.split('/')[0] || `(${files.length} files)`;}

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">{/* --- CONFIGURATION COLUMN --- */}<div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6 space-y-4"><div className="flex items-center justify-between"><h3 className="text-lg font-semibold text-gray-900">1. Upload & Configure</h3><div className="flex items-center space-x-2"><label className="text-sm font-medium text-gray-700">Bulk Mode</label><button onClick={() => setBulkMode(!bulkMode)} className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors ${bulkMode ? 'bg-blue-600' : 'bg-gray-200'}`}><span className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform ${bulkMode ? 'translate-x-6' : 'translate-x-1'}`}/></button></div></div>{!files ? (<label htmlFor="md-upload" className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-lg cursor-pointer bg-gray-50 hover:bg-gray-100"><div className="flex flex-col items-center justify-center">{bulkMode ? <Folder className="w-10 h-10 text-gray-400" /> : <File className="w-10 h-10 text-gray-400" />}<p className="text-sm text-gray-500 mt-2"><span className="font-semibold">Click to select {bulkMode ? 'folder' : 'file'}</span> or drag & drop</p></div><input id="md-upload" type="file" className="hidden" accept=".md,.markdown" onChange={handleFileChange} ref={fileInputRef} {...(bulkMode ? { webkitdirectory: "true", mozdirectory: "true" } : { multiple: true })}/></label>) : (<div className="flex items-center justify-between p-3 bg-green-50 rounded-lg border border-green-200"><div className="flex items-center space-x-3">{bulkMode ? <Folder className="w-6 h-6 text-green-600" /> : <FileText className="w-6 h-6 text-green-600" />}<div><p className="text-sm font-medium text-gray-900">{bulkMode ? getFolderName() : files[0].name}</p><p className="text-xs text-green-700">{files.length} item(s) ready for processing.</p></div></div><button onClick={clearFiles} className="p-1 hover:bg-gray-200 rounded-full transition-colors"><X className="w-4 h-4 text-gray-500" /></button></div>)}<h3 className="text-lg font-semibold text-gray-900 pt-2">2. Prompt</h3><div className="flex space-x-4"><label className="flex items-center space-x-2"><input type="radio" name="promptOption" value="default" checked={promptOption === 'default'} onChange={() => setPromptOption('default')} className="h-4 w-4 text-blue-600"/><span>Default Prompt</span></label><label className="flex items-center space-x-2"><input type="radio" name="promptOption" value="custom" checked={promptOption === 'custom'} onChange={() => setPromptOption('custom')} className="h-4 w-4 text-blue-600"/><span>Custom Prompt</span></label></div>{promptOption === 'custom' ? <textarea value={customPrompt} onChange={(e) => setCustomPrompt(e.target.value)} rows={6} className="w-full p-2 border rounded-md text-sm" /> : <div className="p-3 bg-gray-50 border rounded-md text-xs text-gray-600 whitespace-pre-wrap font-mono">{DEFAULT_PROMPT}</div>}</div>{/* --- ACTION & RESULTS COLUMN --- */}<div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6"><h3 className="text-lg font-semibold text-gray-900 mb-4">3. Process & Review</h3><button onClick={handleProcess} disabled={!files || isLoading} className="w-full flex items-center justify-center space-x-2 px-6 py-3 rounded-lg font-medium transition-all bg-blue-900 text-white hover:bg-blue-800 disabled:bg-gray-300 disabled:cursor-not-allowed">{isLoading ? <Loader className="w-5 h-5 animate-spin" /> : <Sparkles className="w-5 h-5" />}<span>{isLoading ? `Processing ${files?.length} file(s)...` : `Run Inference on ${files?.length || 0} file(s)`}</span></button>{error && <div className="mt-4 p-4 bg-red-50 text-red-800 border border-red-200 rounded-lg flex items-start space-x-2"><AlertTriangle className="w-5 h-5 mt-0.5"/><div><h4 className="font-semibold">Error</h4><p className="text-sm">{error}</p></div></div>}</div></div>
      
      {/* --- RESULTS PREVIEW AREA --- */}
      {bulkResults.length > 0 && (<div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6"><h3 className="text-lg font-semibold text-gray-900 mb-4">Processing Results</h3><div className="flex flex-col lg:flex-row"><div className="w-full lg:w-1/3 lg:pr-6 border-b lg:border-b-0 lg:border-r border-gray-200 mb-4 lg:mb-0"><ul className="space-y-2 max-h-96 overflow-y-auto pr-2">{bulkResults.map((item) => (<li key={item.filename}><div className={`w-full text-left p-2 rounded-lg flex items-center justify-between transition-colors ${selectedResult?.filename === item.filename ? 'bg-blue-100' : 'hover:bg-gray-100'}`}><button onClick={() => setSelectedResult(item)} className="flex items-center space-x-3 flex-grow truncate">{item.result.success ? <CheckCircle className="w-5 h-5 text-green-500 flex-shrink-0" /> : <AlertTriangle className="w-5 h-5 text-red-500 flex-shrink-0" />}<span className="text-sm font-medium text-gray-800 truncate">{item.filename}</span></button><button onClick={(e) => { e.stopPropagation(); handleDownloadSingle(item); }} disabled={!item.result.success} className="p-2 rounded-md hover:bg-gray-200 disabled:text-gray-300 disabled:hover:bg-transparent" aria-label={`Download ${item.filename}`}>
                        <Download className="w-4 h-4 text-gray-600" title="Download as .docx" />
                      </button></div></li>))}</ul></div>{renderSelectedResult()}</div></div>)}
    </div>
  );
};

export default InferenceEngine;