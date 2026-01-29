import { useState, useEffect } from 'react';
import './FrmAttend1.css';
import { Search, Save, Plus, X, Calendar, User, Users, CheckCircle, AlertCircle, Info, FileText } from 'lucide-react';
import sqlQueries from '../sql/FrmAttend1.json';

export default function FrmAttend1() {
  const [activeTab, setActiveTab] = useState(0);
  const [dept, setDept] = useState('All');
  const [part, setPart] = useState('All');
  const [checkIn, setCheckIn] = useState(false);
  const [isWeeklyView, setIsWeeklyView] = useState(false);

  const [mainGridData, setMainGridData] = useState([]);
  const [subGridData, setSubGridData] = useState([]);
  const [deptList, setDeptList] = useState([]);
  const [dbStatus, setDbStatus] = useState('checking'); // checking, connected, disconnected

  useEffect(() => {
    // Check DB Connection
    const checkDb = async () => {
      try {
        const res = await fetch('/api/health');
        const data = await res.json();
        setDbStatus(data.db === 'connected' ? 'connected' : 'disconnected');
      } catch (e) {
        setDbStatus('disconnected');
      }
    };
    checkDb();

    // 1. Fetch Department List on Mount
    const fetchDepartments = async () => {
      // ... existing fetchDepartments code ...
      try {
        const res = await fetch('/api/query', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ queryId: 'm_GetDeptList' })
        });
        const result = await res.json();
        if (result.success) {
          setDeptList(result.rows);
        }
      } catch (err) {
        console.error("Failed to load departments", err);
      }
    };
    fetchDepartments();
  }, []);
  // 2. Fetch Grid Data when Dept/Date changes
  useEffect(() => {
    const fetchData = async () => {
      try {
        console.log("Fetching Main List...");
        // Call Backend API
        const response1 = await fetch('/api/query', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            queryId: 'm_GetMasterList',
            params: ['2023-10-11', dept === 'All' ? '101' : dept]
          })
        });
        const result1 = await response1.json();

        if (result1.success) {
          setMainGridData(result1.rows);
        } else {
          console.error("API Error (MainList):", result1.error);
          setMainGridData([]);
        }

        console.log("Fetching Sub List...");
        const response2 = await fetch('/api/query', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            queryId: 'm_GetSubList',
            params: {
              yyyymm: '202310',
              dept: dept === 'All' ? '101' : dept
            }
          })
        });
        const result2 = await response2.json();
        if (result2.success) {
          setSubGridData(result2.rows);
        }
      } catch (error) {
        console.error("Fetch Error:", error);
      }
    };

    fetchData();
  }, [dept]);

  return (
    <div className="frm-attend-layout">
      <div className="card attend-header">
        <div style={{ position: 'absolute', top: '10px', right: '10px', display: 'flex', alignItems: 'center', gap: '5px', fontSize: '12px' }}>
          <div style={{
            width: '8px',
            height: '8px',
            borderRadius: '50%',
            backgroundColor: dbStatus === 'connected' ? '#10b981' : '#ef4444'
          }} />
          <span style={{ color: dbStatus === 'connected' ? '#065f46' : '#991b1b' }}>
            DB: {dbStatus === 'connected' ? 'Online' : 'Offline'}
          </span>
        </div>
        <div className="filter-group">
          <div className="control-item">
            <label className="lbl">Date</label>
            <div className="input-with-icon">
              <input type="date" defaultValue="2023-10-11" className="input-modern" />
            </div>
          </div>

          <div className="control-item">
            <label className="lbl">Department</label>
            <select value={dept} onChange={(e) => setDept(e.target.value)} className="input-modern">
              <option value="All">Select Department</option>
              {deptList.map((d) => (
                <option key={d.DEPT} value={d.DEPT}>
                  {d.DEPTNAME}
                </option>
              ))}
            </select>
          </div>

          <div className="control-item">
            <label className="lbl">Part</label>
            <select value={part} onChange={(e) => setPart(e.target.value)} className="input-modern">
              <option>All Parts</option>
              <option>Dev</option>
              <option>Design</option>
            </select>
          </div>

          <div className="control-item" style={{ justifyContent: 'flex-end' }}>
            <label className="checkbox-wrapper">
              <input type="checkbox" checked={checkIn} onChange={(e) => setCheckIn(e.target.checked)} />
              <span>Check In Mode</span>
            </label>
            <label className="checkbox-wrapper">
              <input type="checkbox" checked={isWeeklyView} onChange={(e) => setIsWeeklyView(e.target.checked)} />
              <span>Weekly View</span>
            </label>
          </div>
        </div>

        <div className="action-group">
          <button className="btn-modern btn-primary"><Search /> Query</button>
          <button className="btn-modern btn-secondary"><Save /> Save</button>
          <button className="btn-modern btn-secondary"><Plus /> Add</button>
          <button className="btn-modern btn-secondary" style={{ color: '#ef4444', borderColor: '#fee2e2' }}><X /> Close</button>
        </div>
      </div>

      {/* 2. Main Grid */}
      <div className="card grid-section">
        <div className="table-container">
          <table className="modern-table">
            <thead>
              <tr>
                <th style={{ width: '50px' }}>#</th>
                <th>Department</th>
                <th>ID</th>
                <th>Name</th>
                <th>Rank</th>
                <th>Status</th>
                <th>Work Time</th>
                <th>Over Time</th>
                <th>Note</th>
              </tr>
            </thead>
            <tbody>
              {mainGridData.map((row, i) => (
                <tr key={i}>
                  <td style={{ textAlign: 'center', color: '#6b7280' }}>{i + 1}</td>
                  <td>{row.DEPTNAME}</td>
                  <td>{row.SABUN}</td>
                  <td style={{ fontWeight: 500 }}>{row.NAMEK}</td>
                  <td>{row.JIKWI}</td>
                  <td>
                    <span className={`status-badge ${row.STATUS === '0' ? 'status-present' : 'status-absence'}`}>
                      {row.STATUS === '0' ? '재직' : row.STATUS}
                    </span>
                  </td>
                  <td>{row.WORKTIME}</td>
                  <td>{row.OVERTIME}</td>
                  <td>{row.STATUS === '0' ? '' : 'Status: ' + row.STATUS}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* 3. Info Panel (Modernized) */}
      <div className="card info-section">
        <div className="info-item">
          <AlertCircle className="w-4 h-4 icon-danger" />
          <span style={{ color: '#991b1b', fontWeight: 500 }}>Corrective Actions Required (Date & Reason):</span>
          <span style={{ color: '#4b5563' }}>Absence, Sick Leave, Lateness, Early Leave, Sub Work, Public/Special Holiday.</span>
        </div>
        <div className="info-item">
          <Info className="w-4 h-4 icon-success" />
          <span style={{ color: '#065f46' }}>Business Trips must include detailed reports. Saturday non-work: Check 'Non-work'.</span>
        </div>
        <div className="info-item">
          <AlertCircle className="w-4 h-4 icon-warning" />
          <span style={{ color: '#92400e', fontWeight: 600 }}>COVID-19 Advisory:</span>
          <span style={{ color: '#92400e' }}>Contact Infection Control immediately if symptoms appear (Cough, Fever, etc.).</span>
        </div>
      </div>

      {/* 4. Bottom Tabs & Grid */}
      <div className="card bottom-section">
        <div className="tabs-header">
          <div className={`tab-link ${activeTab === 0 ? 'active' : ''}`} onClick={() => setActiveTab(0)}>
            Weekly Work Time
          </div>
          <div className={`tab-link ${activeTab === 1 ? 'active' : ''}`} onClick={() => setActiveTab(1)}>
            Short Time Records
          </div>
        </div>
        <div className="tab-body">
          <div className="table-container" style={{ flex: 1 }}>
            <table className="modern-table">
              <thead>
                <tr>
                  <th style={{ width: '50px' }}>#</th>
                  <th>Department</th>
                  <th>ID</th>
                  <th>Name</th>
                  <th>Status</th>
                  <th>In Time</th>
                  <th>Out Time</th>
                  <th>Total Hours</th>
                  <th>Notes</th>
                </tr>
              </thead>
              <tbody>
                {subGridData.map((row, i) => (
                  <tr key={i}>
                    <td style={{ textAlign: 'center', color: '#6b7280' }}>{i + 1}</td>
                    <td>{row.dept}</td>
                    <td>{row.id}</td>
                    <td style={{ fontWeight: 500 }}>{row.name}</td>
                    <td>
                      <span className={`status-badge ${row.status === 'Present' ? 'status-present' : 'status-absence'}`}>
                        {row.status}
                      </span>
                    </td>
                    <td>{row.inTime}</td>
                    <td>{row.outTime}</td>
                    <td style={{ fontWeight: 600, color: '#4f46e5' }}>{row.workTime}h</td>
                    <td>{row.reason}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="side-actions">
            <button className="action-btn">
              <FileText className="w-5 h-5 mb-1" />
              <span>View<br />Ledger</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
