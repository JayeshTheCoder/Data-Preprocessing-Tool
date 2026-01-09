// AppContext.tsx

import React, { createContext, useState, useContext, ReactNode } from 'react';

interface UploadedFile {
  id: string;
  name: string;
  size: number;
  status: 'uploaded' | 'processing' | 'cleaned' | 'error';
  type: string;
  file: File;
}

interface AppContextType {
  uploadedFiles: UploadedFile[];
  setUploadedFiles: React.Dispatch<React.SetStateAction<UploadedFile[]>>;
  sessionId: string | null;
  setSessionId: React.Dispatch<React.SetStateAction<string | null>>;
  bulkMode: boolean;
  setBulkMode: React.Dispatch<React.SetStateAction<boolean>>;
  selectedMetric: string;
  setSelectedMetric: React.Dispatch<React.SetStateAction<string>>;
  expandedMetrics: string[];
  toggleMetric: (metricId: string) => void;
  selectedSubMetric: string | null;
  setSelectedSubMetric: React.Dispatch<React.SetStateAction<string | null>>;
  selectedWcType: 'dso' | 'overhead';
  setSelectedWcType: (value: 'dso' | 'overhead') => void;
  selectedRules: { [key: string]: boolean };
  setSelectedRules: React.Dispatch<React.SetStateAction<{ [key: string]: boolean }>>;
  vendorAnalysisType: 'mom' | 'qtd';
  setVendorAnalysisType: React.Dispatch<React.SetStateAction<'mom' | 'qtd'>>;
  
}

const AppContext = createContext<AppContextType | undefined>(undefined);

export const AppProvider = ({ children }: { children: ReactNode }) => {
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [sessionId, setSessionId] = useState<string | null>(null);
  const [bulkMode, setBulkMode] = useState(false);
  const [selectedMetric, setSelectedMetric] = useState('sales');
  const [expandedMetrics, setExpandedMetrics] = useState<string[]>(['data-processing']);
  const [selectedSubMetric, setSelectedSubMetric] = useState<string | null>('pex-bi');
  const [selectedRules, setSelectedRules] = useState({
    removeDuplicates: true,
    groupUnits: true, 
    validateFormats: true,
    standardizeNames: false,
    removeOutliers: false,
    normalizeData: false
  });
  const [vendorAnalysisType, setVendorAnalysisType] = useState<'mom' | 'qtd'>('mom');
  
  const toggleMetric = (metricId: string) => {
    setExpandedMetrics(prev =>
      prev.includes(metricId)
        ? prev.filter(id => id !== metricId)
        : [...prev, metricId]
    );
  };

  // --- COMPATIBILITY SHIM ---
  const selectedWcType = (
    selectedSubMetric === 'dso' || selectedSubMetric === 'overhead' 
    ? selectedSubMetric 
    : 'dso'
  ) as 'dso' | 'overhead';

  const setSelectedWcType = (value: 'dso' | 'overhead') => {
    setSelectedSubMetric(value);
  };
  // --- END COMPATIBILITY SHIM ---

  return (
    <AppContext.Provider
      value={{
        uploadedFiles,
        setUploadedFiles,
        sessionId,
        setSessionId,
        bulkMode,
        setBulkMode,
        selectedMetric,
        setSelectedMetric,
        expandedMetrics,
        toggleMetric,
        selectedSubMetric,
        setSelectedSubMetric,
        selectedWcType,
        setSelectedWcType,
        selectedRules,
        setSelectedRules,
        vendorAnalysisType,
        setVendorAnalysisType
      }}
    >
      {children}
    </AppContext.Provider>
  );
};

export const useApp = () => {
  const context = useContext(AppContext);
  if (context === undefined) {
    throw new Error('useApp must be used within an AppProvider');
  }
  return context;
};