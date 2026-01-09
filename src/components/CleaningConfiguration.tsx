import React from 'react'; // No longer needs useState
import { Settings, Filter, Zap, Shield, Users } from 'lucide-react';
import { useApp } from '../context/AppContext'; // Import the hook

const CleaningConfiguration = () => {
  // --- KEY CHANGE ---
  // Get state and setter from the context, not local useState
  const { bulkMode, setBulkMode, selectedRules, setSelectedRules } = useApp();
  // --- END KEY CHANGE ---

  const toggleRule = (rule: keyof typeof selectedRules) => {
    // This now updates the *global* state in the context
    setSelectedRules(prev => ({
      ...prev,
      [rule]: !prev[rule]
    }));
  };

  const cleaningRules = [
    { key: 'removeDuplicates', label: 'Remove Duplicates', icon: Filter, description: 'Remove duplicate rows based on key fields' },
    { key: 'groupUnits', label: 'Group Units', icon: Users, description: 'Merge outputs by "Grouping Unit" from the directory file' },
    { key: 'validateFormats', label: 'Validate Using Hyperion', icon: Zap, description: 'Ensure the data is correct in accordance with hyperion' },
    { key: 'standardizeNames', label: 'Standardize Names', icon: Settings, description: 'Normalize naming conventions' },
    { key: 'removeOutliers', label: 'Remove Outliers', icon: Filter, description: 'Identify and handle statistical outliers' },
    { key: 'normalizeData', label: 'Normalize Data', icon: Zap, description: 'Scale numerical values to standard ranges' }
  ];
  
  // No other changes are needed. The rest of your file is perfect.
  // The UI will read from 'selectedRules' from the context.
  
  return (
    <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
      <div className="flex items-center justify-between mb-4">
        <h3 className="text-lg font-semibold text-gray-900">Cleaning Configuration</h3>
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

      <div className="space-y-4">
        <div className="grid grid-cols-1 gap-3">
          {cleaningRules.map((rule) => {
            const Icon = rule.icon;
            // This 'isSelected' now correctly reads from the global context
            const isSelected = selectedRules[rule.key as keyof typeof selectedRules];
            
            return (
              <div
                key={rule.key}
                className={`border rounded-lg p-4 cursor-pointer transition-all ${
                  isSelected 
                    ? 'border-blue-200 bg-blue-50' 
                    : 'border-gray-200 hover:border-gray-300'
                }`}
                onClick={() => toggleRule(rule.key as keyof typeof selectedRules)}
              >
                <div className="flex items-start space-x-3">
                  <div className={`mt-1 ${isSelected ? 'text-blue-600' : 'text-gray-400'}`}>
                    <Icon className="w-5 h-5" />
                  </div>
                  <div className="flex-1">
                    <div className="flex items-center space-x-2">
                      <h4 className={`font-medium ${isSelected ? 'text-blue-900' : 'text-gray-900'}`}>
                        {rule.label}
                      </h4>
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={() => toggleRule(rule.key as keyof typeof selectedRules)}
                        className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded"
                      />
                    </div>
                    <p className="text-sm text-gray-600 mt-1">{rule.description}</p>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
        {/* ... rest of your component ... */}
      </div>
    </div>
  );
};

export default CleaningConfiguration;