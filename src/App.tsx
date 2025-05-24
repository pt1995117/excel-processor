import React from 'react';
import { ConfigProvider } from 'antd';
import zhCN from 'antd/es/locale/zh_CN';
import ExcelProcessor from './components/ExcelProcessor';
import 'antd/dist/antd.min.css';
import './App.css';

const App: React.FC = () => {
  return (
    <ConfigProvider locale={zhCN}>
      <div style={{ 
        minHeight: '100vh',
        background: 'linear-gradient(135deg, #f5f7fa 0%, #e4e8eb 100%)',
        padding: '40px 20px'
      }}>
        <div style={{ 
          maxWidth: '1200px', 
          margin: '0 auto',
          backgroundColor: '#fff',
          borderRadius: '16px',
          boxShadow: '0 4px 20px rgba(0, 0, 0, 0.08)',
          overflow: 'hidden'
        }}>
          <ExcelProcessor />
        </div>
      </div>
    </ConfigProvider>
  );
};

export default App;
