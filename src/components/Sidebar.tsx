// Sidebar.tsx

import React from 'react';
import { 
  ChevronLeft, 
  ChevronRight, 
  TrendingUp, 
  DollarSign, 
  Activity, 
  BarChart, 
  Briefcase, 
  FileText,
  BarChart3 // <-- Import new icon
} from 'lucide-react';
import { useApp } from '../context/AppContext';

interface SidebarProps {
  collapsed: boolean;
  onToggle: () => void;
}

// --- UPDATED METRICS ARRAY ---
const metrics = [
  { id: 'sales', name: 'Sales', icon: TrendingUp, color: 'text-green-600', units: [] },
  { 
    id: 'pex', 
    name: 'PEX', 
    icon: DollarSign, 
    color: 'text-blue-600', 
    units: [
      { id: 'pex-bi', name: 'PEX BI Data' },
      { id: 'pex-vendor', name: 'PEX Vendor Data' },
    ] 
  },
  // { id: 'c1', name: 'C1', icon: Activity, color: 'text-purple-600', units: [] },
  { id: 'oe', name: 'OE', icon: BarChart, color: 'text-orange-600', units: [] }, 
  { 
    id: 'working-capital', 
    name: 'Working Capital', 
    icon: Briefcase, 
    color: 'text-indigo-600', 
    units: [
      { id: 'dso', name: 'DSO', icon: BarChart3 },
      { id: 'overhead', name: 'ITO (Overhead)', icon: TrendingUp }
    ] 
  },
  { id: 'inference', name: 'Controller Refinery', icon: FileText, color: 'text-slate-600', units: [] }
];

const Sidebar: React.FC<SidebarProps> = ({ collapsed, onToggle }) => {
  // Assume selectedSubMetric and setSelectedSubMetric are added to your AppContext
  const { 
    selectedMetric, 
    setSelectedMetric, 
    expandedMetrics, 
    toggleMetric,
    selectedSubMetric,    // <-- NEW: Get sub-metric state from context
    setSelectedSubMetric  // <-- NEW: Get setter for sub-metric from context
  } = useApp();

  return (
    <aside className={`bg-white border-r border-gray-200 transition-all duration-300 ${collapsed ? 'w-16' : 'w-80'}`}>
      <div className="p-4 border-b border-gray-200 flex justify-between items-center">
        {!collapsed && <h2 className="font-semibold text-gray-800">Metrics</h2>}
        <button 
          onClick={onToggle}
          className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
        >
          {collapsed ? <ChevronRight className="w-4 h-4" /> : <ChevronLeft className="w-4 h-4" />}
        </button>
      </div>
      
      <div className="p-4 space-y-2">
        {metrics.map((metric) => {
          const Icon = metric.icon;
          const isSelected = selectedMetric === metric.id;
          const isExpanded = expandedMetrics.includes(metric.id);
          
          return (
            <div key={metric.id} className="space-y-1">
              <button
                onClick={() => {
                  setSelectedMetric(metric.id);
                  // Reset sub-metric when changing main metric
                  if(metric.units.length > 0) {
                    setSelectedSubMetric(metric.units[0].id);
                  } else {
                    setSelectedSubMetric(null);
                  }
                  if (!collapsed) toggleMetric(metric.id);
                }}
                className={`w-full flex items-center space-x-3 p-3 rounded-lg transition-all ${
                  isSelected 
                    ? 'bg-blue-50 border border-blue-200 text-blue-800' 
                    : 'hover:bg-gray-50 text-gray-700'
                }`}
              >
                <Icon className={`w-5 h-5 ${isSelected ? 'text-blue-600' : metric.color}`} />
                {!collapsed && (
                  <span className="flex-1 text-left font-medium">{metric.name}</span>
                )}
              </button>
              
              {/* --- NEW: RENDER SUB-METRIC OPTIONS --- */}
              {!collapsed && isExpanded && metric.units.length > 0 && (
                <div className="pl-8 pr-2 pt-1 pb-2 space-y-1">
                  {metric.units.map((unit) => {
                    const isSubSelected = selectedSubMetric === unit.id;
                    const UnitIcon = unit.icon;
                    return (
                       <button
                        key={unit.id}
                        onClick={() => setSelectedSubMetric(unit.id)}
                        className={`w-full flex items-center text-left p-2 text-sm rounded-md transition-colors ${
                          isSubSelected 
                          ? 'bg-indigo-100 font-semibold text-indigo-800' 
                          : 'text-gray-600 hover:bg-gray-100'
                        }`}
                      >
                        {UnitIcon && <UnitIcon className={`w-4 h-4 mr-3 ${isSubSelected ? 'text-indigo-600' : 'text-gray-500'}`} />}
                        <span>{unit.name}</span>
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </aside>
  );
};

export default Sidebar;