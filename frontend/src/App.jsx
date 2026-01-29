import { useState } from 'react';
import FrmLogin from './forms/frmLogin';
import FrmPatient from './forms/frmPatient';
import FrmAttend1 from './forms/FrmAttend1';
import './App.css';

function App() {
  const [currentScreen, setCurrentScreen] = useState('attend'); // Default to attend for dev

  const handleLoginSuccess = () => {
    setCurrentScreen('patient');
  };

  return (
    <div>
      {/* Development Navigation Header */}
      <div style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        zIndex: 1000,
        background: '#333',
        padding: '8px',
        display: 'flex',
        gap: '10px',
        justifyContent: 'center'
      }}>
        <button onClick={() => setCurrentScreen('login')} style={{ padding: '4px 8px' }}>Login</button>
        <button onClick={() => setCurrentScreen('patient')} style={{ padding: '4px 8px' }}>Patient</button>
        <button onClick={() => setCurrentScreen('attend')} style={{ padding: '4px 8px' }}>Attend Check</button>
      </div>

      <div style={{ marginTop: '50px', height: 'calc(100vh - 50px)' }}>
        {currentScreen === 'login' && <FrmLogin onLoginSuccess={handleLoginSuccess} />}
        {currentScreen === 'patient' && <FrmPatient />}
        {currentScreen === 'attend' && <FrmAttend1 />}
      </div>
    </div>
  );
}

export default App;
