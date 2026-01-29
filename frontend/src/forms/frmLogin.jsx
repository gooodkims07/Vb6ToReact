
import { useState, useEffect, useRef } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { User, Lock, LogIn, X, Key } from 'lucide-react';
import './frmLogin.css';

export default function FrmLogin({ onLoginSuccess }) {
    const [txtID, setTxtID] = useState("");
    const [txtPassword, setTxtPassword] = useState("");
    const [errorID, setErrorID] = useState(false);
    const [errorPass, setErrorPass] = useState(false);

    const txtIDRef = useRef(null);
    const txtPasswordRef = useRef(null);

    useEffect(() => {
        setTxtID("");
        setTxtPassword("");
        if (txtIDRef.current) txtIDRef.current.focus();
    }, []);

    const handleConfirmClick = async () => {
        setErrorID(false);
        setErrorPass(false);

        if (txtID.trim() === "") {
            setErrorID(true);
            if (txtIDRef.current) txtIDRef.current.focus();
            return;
        }

        if (txtPassword.trim() === "") {
            setErrorPass(true);
            if (txtPasswordRef.current) txtPasswordRef.current.focus();
            return;
        }

        try {
            const response = await fetch('/api/query', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    queryId: 'm_LoginCheck',
                    params: {
                        sabun: txtID,
                        password: txtPassword
                    }
                })
            });

            const data = await response.json();

            if (data.success && data.rows && data.rows.length > 0) {
                // Login Success
                onLoginSuccess(data.rows[0]);
            } else {
                alert("아이디 또는 패스워드가 틀립니다.");
                setTxtID("");
                setTxtPassword("");
                if (txtIDRef.current) txtIDRef.current.focus();
            }
        } catch (err) {
            console.error("Login error:", err);
            alert("로그인 처리 중 오류가 발생했습니다: " + err.message);
        }
    };

    const handleExitClick = () => {
        if (confirm("프로그램을 종료하시겠습니까?")) {
            window.close();
        }
    };

    const [isPwChangeModalOpen, setIsPwChangeModalOpen] = useState(false);
    const [oldPw, setOldPw] = useState("");
    const [newPw, setNewPw] = useState("");
    const [confirmPw, setConfirmPw] = useState("");
    const [changePwError, setChangePwError] = useState("");

    const handlePwChangeClick = () => {
        setIsPwChangeModalOpen(true);
        setOldPw("");
        setNewPw("");
        setConfirmPw("");
        setChangePwError("");
    };

    const handlePwChangeSubmit = async () => {
        setChangePwError("");

        if (!txtID) {
            setChangePwError("아이디를 입력해주세요.");
            return;
        }

        if (newPw !== confirmPw) {
            setChangePwError("새 암호가 일치하지 않습니다.");
            return;
        }

        if (newPw.length < 4) {
            setChangePwError("새 암호는 4자리 이상이어야 합니다.");
            return;
        }

        try {
            const response = await fetch('/api/query', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    queryId: 'm_UpdatePassword',
                    params: {
                        sabun: txtID,
                        oldPassword: oldPw,
                        newPassword: newPw
                    }
                })
            });

            const data = await response.json();

            if (data.success && data.rows && data.rows.rowsAffected > 0) {
                alert("암호가 성공적으로 변경되었습니다.");
                setIsPwChangeModalOpen(false);
            } else if (data.success && data.rows && data.rows.rowsAffected === 0) {
                setChangePwError("기존 암호가 틀리거나 사용자를 찾을 수 없습니다.");
            } else {
                setChangePwError("암호 변경 실패: " + (data.error || "알 수 없는 오류"));
            }
        } catch (err) {
            setChangePwError("서버 오류: " + err.message);
        }
    };

    const [txtName, setTxtName] = useState("");

    const handleIdBlur = async () => {
        if (!txtID || txtID.trim() === "") {
            setTxtName("");
            return;
        }

        try {
            const response = await fetch('/api/query', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    queryId: 'm_GetName',
                    params: { sabun: txtID }
                })
            });

            const data = await response.json();
            if (data.success && data.rows && data.rows.length > 0) {
                setTxtName(data.rows[0].NAME);
            } else {
                setTxtName(""); // Clear if not found
            }
        } catch (err) {
            console.error(err);
            // Silent error or name clear
        }
    };

    return (
        <motion.div
            initial={{ opacity: 0, scale: 0.9 }}
            animate={{ opacity: 1, scale: 1 }}
            className="frmLogin-container"
        >
            <div className="frmLogin-header">
                <div className="frmLogin-title">환영합니다!!</div>
                <div className="frmLogin-subtitle">KYcare 로그인하세요</div>
            </div>

            <div className="input-group">
                <label className="input-label">아이디</label>
                <div style={{ display: 'flex', gap: '8px' }}>
                    <div className="input-wrapper" style={{ flex: 1 }}>
                        <User className="input-icon" size={18} />
                        <input
                            type="text"
                            className={`modern-input ${errorID ? 'error' : ''}`}
                            placeholder="Enter ID"
                            value={txtID}
                            onChange={(e) => { setTxtID(e.target.value); setErrorID(false); }}
                            onBlur={handleIdBlur}
                            ref={txtIDRef}
                        />
                    </div>

                    {/* Name Display */}
                    <div className="input-wrapper" style={{ flex: 1 }}>
                        <input
                            type="text"
                            className="modern-input"
                            style={{ backgroundColor: '#f3f4f6', color: '#666', border: 'none' }}
                            placeholder="Name"
                            value={txtName}
                            readOnly
                        />
                    </div>
                </div>
            </div>

            <div className="input-group">
                <label className="input-label">암호</label>
                <div className="input-wrapper">
                    <Lock className="input-icon" size={18} />
                    <input
                        type="password"
                        className={`modern-input ${errorPass ? 'error' : ''}`}
                        placeholder="Enter your password"
                        value={txtPassword}
                        onChange={(e) => { setTxtPassword(e.target.value); setErrorPass(false); }}
                        ref={txtPasswordRef}
                        onKeyDown={(e) => {
                            if (e.key === 'Enter') handleConfirmClick();
                        }}
                    />
                </div>
            </div>

            <div className="button-group">
                <motion.button
                    whileHover={{ scale: 1.02 }}
                    whileTap={{ scale: 0.98 }}
                    className="modern-btn btn-primary"
                    onClick={handleConfirmClick}
                >
                    <LogIn size={18} />
                    로그인
                </motion.button>

                <motion.button
                    whileHover={{ scale: 1.02 }}
                    whileTap={{ scale: 0.98 }}
                    className="modern-btn btn-secondary btn-wide"
                    onClick={handlePwChangeClick}
                >
                    <Key size={18} />
                    암호변경
                </motion.button>

                <motion.button
                    whileHover={{ scale: 1.02 }}
                    whileTap={{ scale: 0.98 }}
                    className="modern-btn btn-secondary"
                    onClick={handleExitClick}
                >
                    <X size={18} />
                    종료
                </motion.button>
            </div>

            <AnimatePresence>
                {isPwChangeModalOpen && (
                    <motion.div
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        exit={{ opacity: 0 }}
                        className="modal-overlay"
                    >
                        <motion.div
                            initial={{ scale: 0.9, y: 20 }}
                            animate={{ scale: 1, y: 0 }}
                            exit={{ scale: 0.9, y: 20 }}
                            className="modal-content"
                        >
                            <div className="modal-header">
                                <div className="modal-title">비밀번호 변경</div>
                                <button className="close-btn" onClick={() => setIsPwChangeModalOpen(false)}>
                                    <X size={20} />
                                </button>
                            </div>

                            <div className="modal-body">
                                <div className="input-group">
                                    <label className="input-label">기존 암호</label>
                                    <div className="input-wrapper">
                                        <Lock className="input-icon" size={18} />
                                        <input
                                            type="password"
                                            className="modern-input"
                                            placeholder="Old Password"
                                            value={oldPw}
                                            onChange={(e) => setOldPw(e.target.value)}
                                        />
                                    </div>
                                </div>

                                <div className="input-group">
                                    <label className="input-label">새 암호</label>
                                    <div className="input-wrapper">
                                        <Key className="input-icon" size={18} />
                                        <input
                                            type="password"
                                            className="modern-input"
                                            placeholder="New Password"
                                            value={newPw}
                                            onChange={(e) => setNewPw(e.target.value)}
                                        />
                                    </div>
                                </div>

                                <div className="input-group">
                                    <label className="input-label">새 암호 확인</label>
                                    <div className="input-wrapper">
                                        <Key className="input-icon" size={18} />
                                        <input
                                            type="password"
                                            className="modern-input"
                                            placeholder="Confirm New Password"
                                            value={confirmPw}
                                            onChange={(e) => setConfirmPw(e.target.value)}
                                        />
                                    </div>
                                </div>

                                {changePwError && <div style={{ color: 'red', fontSize: '13px' }}>{changePwError}</div>}

                                <button className="modern-btn btn-primary" onClick={handlePwChangeSubmit}>
                                    변경하기
                                </button>
                            </div>
                        </motion.div>
                    </motion.div>
                )}
            </AnimatePresence>
        </motion.div>
    );
}
