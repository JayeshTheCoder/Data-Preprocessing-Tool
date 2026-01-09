import React from 'react';
import { Home, BarChart3, Settings, Database } from 'lucide-react';
import { useApp } from '../context/AppContext';

const Header = () => {
  const { setSelectedMetric } = useApp();
  return (
    <header className="bg-blue-900 text-white h-16 flex items-center px-6 shadow-lg">
      <div className="flex items-center space-x-3">
        <Database className="w-8 h-8 text-blue-300" />
        <h1 className="text-xl font-bold">DataCleanse Pro</h1>
      </div>
      
      <nav className="ml-12 flex space-x-8">
        <button className="flex items-center space-x-2 px-3 py-2 rounded-lg hover:bg-blue-800 transition-colors">
          <Home className="w-4 h-4" />
          <span>Home</span>
        </button>
        <button onClick={() => setSelectedMetric('processing-pipeline')} className="flex items-center space-x-2 px-3 py-2 rounded-lg hover:bg-blue-800 transition-colors">
          <BarChart3 className="w-4 h-4" />
          <span>Processing Pipeline</span>
        </button>
        
        {/* <button className="flex items-center space-x-2 px-3 py-2 rounded-lg hover:bg-blue-800 transition-colors">
          <Database className="w-4 h-4" />
          <span>Bulk Processing</span>
        </button> */}
        <button className="flex items-center space-x-2 px-3 py-2 rounded-lg hover:bg-blue-800 transition-colors">
          <Settings className="w-4 h-4" />
          <span>Settings</span>
        </button>
      </nav>
    </header>
  );
};

export default Header;