// WorkingCapitalSelector.tsx

import React from 'react';
import { BarChart3, TrendingUp } from 'lucide-react';
import { useApp } from '../context/AppContext';

const wcTypes = [
  { key: 'dso', label: 'DSO', icon: BarChart3, description: 'Days Sales Outstanding Analysis' },
  { key: 'overhead', label: 'ITO (Overhead)', icon: TrendingUp, description: 'Inventory & Overhead Cost Analysis' },
];

const WorkingCapitalSelector = () => {
  const { selectedWcType, setSelectedWcType } = useApp();

  return (
    <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6 mb-6">
      <h3 className="text-lg font-semibold text-gray-900 mb-4">Select Working Capital Analysis Type</h3>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {wcTypes.map((wcType) => (
          <button
            key={wcType.key}
            onClick={() => setSelectedWcType(wcType.key as 'dso' | 'overhead')}
            className={`flex items-center text-left p-4 rounded-lg border-2 transition-all ${
              selectedWcType === wcType.key
                ? 'border-indigo-600 bg-indigo-50 shadow-md'
                : 'border-gray-300 hover:border-gray-400 hover:bg-gray-50'
            }`}
          >
            <wcType.icon
              className={`w-8 h-8 mr-4 ${
                selectedWcType === wcType.key ? 'text-indigo-600' : 'text-gray-400'
              }`}
            />
            <div>
              <p className="font-semibold text-gray-900">{wcType.label}</p>
              <p className="text-sm text-gray-500">{wcType.description}</p>
            </div>
          </button>
        ))}
      </div>
    </div>
  );
};

export default WorkingCapitalSelector;