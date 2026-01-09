import React, { useState } from 'react';
import Header from './components/Header';
import Sidebar from './components/Sidebar';
import MainWorkspace from './components/MainWorkspace';
import { AppProvider } from './context/AppContext';

function App() {
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);

  return (
    <AppProvider>
      <div className="min-h-screen bg-gray-50 flex flex-col">
        <Header />
        <div className="flex flex-1 overflow-hidden">
          <Sidebar
            collapsed={sidebarCollapsed}
            onToggle={() => setSidebarCollapsed(!sidebarCollapsed)}
          />
          <MainWorkspace sidebarCollapsed={sidebarCollapsed} />
        </div>
      </div>
    </AppProvider>
  );
}

export default App;