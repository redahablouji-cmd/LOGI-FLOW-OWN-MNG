/**
 * @license
 * SPDX-License-Identifier: Apache-2.5
 */

import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import LoginPage from './pages/LoginPage';
import RegisterCompanyPage from './pages/RegisterCompanyPage';
import OwnerDashboard from './pages/OwnerDashboard';
import ManagerDashboard from './pages/ManagerDashboard';
import { Toaster } from 'sonner';

export default function App() {
  return (
    <>
      <BrowserRouter>
        <Routes>
          {/* Base Redirection */}
          <Route path="/" element={<Navigate to="/login" replace />} />

          {/* Core Access Portals */}
          <Route path="/login" element={<LoginPage />} />
          <Route path="/register" element={<RegisterCompanyPage />} />

          {/* Role Secured Dashboards */}
          <Route path="/owner" element={<OwnerDashboard />} />
          <Route path="/manager" element={<ManagerDashboard />} />

          {/* Catch All - Redirect to login */}
          <Route path="*" element={<Navigate to="/login" replace />} />
        </Routes>
      </BrowserRouter>
      
      {/* Toast Notification Provider */}
      <Toaster 
        position="top-right" 
        expand={true} 
        richColors 
        closeButton 
        theme="light"
      />
    </>
  );
}

