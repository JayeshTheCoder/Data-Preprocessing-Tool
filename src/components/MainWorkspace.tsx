// MainWorkspace.tsx

import React from 'react';
import FileUploadArea from './FileUploadArea';
// import DataPreview from './DataPreview';
import CleaningConfiguration from './CleaningConfiguration';
import ExecutionControls from './ExecutionControls';
import InferenceEngine from './InferenceEngine';
import ProcessingPipeline from './ProcessingPipeline';
import { useApp } from '../context/AppContext';

interface MainWorkspaceProps {
  sidebarCollapsed: boolean;
}

const MainWorkspace: React.FC<MainWorkspaceProps> = ({ sidebarCollapsed }) => {
  const { selectedMetric } = useApp();

  const getMetricConfig = (metric: string) => {
    const configs: { [key: string]: { name: string, description: string } } = {
      'sales': { name: 'Sales', description: 'Upload, clean, and process your sales data files' },
      'pex': { name: 'PEX', description: 'Upload, clean, and process your PEX data files' },
      'c1': { name: 'C1', description: 'Upload, clean, and process your C1 data files' },
      'oe': { name: 'OE', description: 'Upload, clean, and process your OE data files' },
      'working-capital': { name: 'Working Capital', description: 'Upload, clean, and process your working capital data files' },
      'inference': { name: 'Controller Refinery', description: 'Analyze markdown documents using an AI-powered engine' },
      'processing-pipeline': { name: 'Processing Pipeline', description: 'Processing pipeline tool' }
    };
    return configs[metric] || { name: 'Select Metric', description: 'Please select a metric from the sidebar.' };
  };
  
  const metricConfig = getMetricConfig(selectedMetric);

  const renderDataProcessingWorkspace = () => (
    <>
      <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 mb-6">
        <FileUploadArea />
        <CleaningConfiguration />
      </div>
      {/* <DataPreview /> */}
      <ExecutionControls />
    </>
  );

  return (
    <main className={`flex-1 overflow-auto transition-all duration-300 ${sidebarCollapsed ? 'ml-0' : 'ml-0'}`}>
      <div className="p-6 max-w-7xl mx-auto">
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-gray-900 mb-2">
            {metricConfig.name}
          </h1>
          <p className="text-gray-600">
            {metricConfig.description}
          </p>
        </div>
        
        {/* --- FIX: Corrected the component name from InFERENCEEngine to InferenceEngine --- */}
        {/* Render Processing Pipeline when header selects 'processing-pipeline' */}
        {selectedMetric === 'inference' && <InferenceEngine />}
        {selectedMetric === 'processing-pipeline' && <ProcessingPipeline />}
        {selectedMetric !== 'inference' && selectedMetric !== 'processing-pipeline' && renderDataProcessingWorkspace()}
      </div>
    </main>
  );
};

export default MainWorkspace;