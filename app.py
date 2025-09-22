from flask import Flask, render_template_string, send_file, request, flash, redirect, url_for, session, current_app, get_flashed_messages
from werkzeug.utils import secure_filename
from generate_sale_order import prepare_data, write_report
import os
import logging
from datetime import datetime
import pandas as pd
import uuid
from functools import wraps
import time
import sqlite3
import math
from flask_bcrypt import Bcrypt
import re # Import re for filename sanitization

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler()
    ]
)

# 🔒 Track active user sessions (for single-login restriction)
ACTIVE_SESSIONS = {}

app = Flask(__name__)
# --- SECURE: Load secret key from environment variables ---
app.secret_key = os.getenv('SECRET_KEY')
if not app.secret_key:
    raise ValueError("SECRET_KEY environment variable is not set!")

# Session configuration for better security
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SECURE'] = True  # Set to True in production with HTTPS
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['PERMANENT_SESSION_LIFETIME'] = 3600  # 1 hour session timeout

# Initialize Bcrypt for password hashing
bcrypt = Bcrypt(app)

# Session validation middleware
@app.before_request
def validate_session():
    # Skip validation for static files and login/logout routes
    if request.endpoint in ['static', 'login', 'logout', 'favicon']:
        return
    
    # Skip for non-authenticated routes (add more if needed)
    public_routes = ['login']
    if request.endpoint in public_routes:
        return
    
    # Check if user session exists and is valid
    if 'user' in session:
        username = session['user']
        if username not in ACTIVE_SESSIONS:
            app.logger.warning(f"Invalid session detected for user: {username}")
            session.clear()
            flash("Your session is invalid. Please login again.", "error")
            return redirect(url_for('login'))
        
        if username not in USERS:
            app.logger.error(f"User {username} not found in USERS")
            session.clear()
            ACTIVE_SESSIONS.pop(username, None)
            flash("User account not found. Please contact administrator.", "error")
            return redirect(url_for('login'))

# Load users from .env (format: USER1=username:hashed_password)
USERS = {}
for key, value in os.environ.items():
    if key.startswith("USER") and key not in ("USERNAME","USERDOMAIN","USERDOMAIN_ROAMINGPROFILE"):
        try:
            uname, hashed_pwd = value.split(":", 1)
            USERS[uname] = hashed_pwd
        except ValueError:
            app.logger.warning(f"Invalid user format in {key} (expected USERX=username:hashed_password)")

# === RENDER MODIFICATION: POINT TO PERSISTENT DISK ===
# The mount path for our Render Disk will be '/var/data'
DATA_DIR = '/var/data'

# Configuration
class Config:
    UPLOAD_FOLDER = os.path.join(DATA_DIR, 'uploads')
    REPORT_FOLDER = os.path.join(DATA_DIR, 'reports')
    DATABASE_FILE = os.path.join(DATA_DIR, 'order_counter.db')
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size
    ALLOWED_EXTENSIONS = {'xls', 'xlsx'}
    CLEANUP_INTERVAL_HOURS = int(os.getenv('CLEANUP_INTERVAL_HOURS', 8))

app.config.from_object(Config)
# =======================================================

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['REPORT_FOLDER'], exist_ok=True)

# ---------------- Enhanced HTML Templates with Modern UI/UX ----------------
THEME_CSS_AND_JS = """
<style>
    /* Modern CSS Reset & Base */
    *, *::before, *::after { 
        margin: 0; padding: 0; box-sizing: border-box; 
    }
    
    /* PREMIUM THEME-SWITCHER STYLES */
    .header-controls {
        position: absolute; top: 1.5rem; right: 2rem; display: flex; align-items: center; 
        gap: 1rem; z-index: 100;
    }
    .theme-selector select {
        background: rgba(255, 255, 255, 0.15); backdrop-filter: blur(20px); 
        color: var(--text-light, #fff); border: 1px solid rgba(255, 255, 255, 0.25);
        border-radius: 50px; padding: 0.6rem 1rem; font-size: 0.85rem; cursor: pointer;
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1); font-weight: 500;
        -webkit-appearance: none; -moz-appearance: none; appearance: none; 
        padding-right: 3rem; min-width: 140px;
        background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23ffffff' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3e%3cpolyline points='6 9 12 15 18 9'%3e%3c/polyline%3e%3c/svg%3e");
        background-repeat: no-repeat; background-position: right 1rem center; background-size: 1rem;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
    }
    .theme-selector select:hover { 
        background: rgba(255, 255, 255, 0.25); transform: translateY(-2px); 
        box-shadow: 0 12px 40px rgba(0, 0, 0, 0.15);
    }
    .theme-selector select:focus { 
        outline: none; border-color: rgba(255, 255, 255, 0.5); 
        box-shadow: 0 0 0 4px rgba(255, 255, 255, 0.1);
    }
    .theme-selector option { 
        background-color: #1a202c; color: #f7fafc; padding: 0.5rem;
        border-radius: 8px; margin: 0.25rem 0;
    }
    .login-container .theme-selector, .card .theme-selector {
        position: absolute; top: 2rem; right: 2rem;
    }
    .login-container .theme-selector select, .card .theme-selector select {
        background: rgba(0, 0, 0, 0.05); color: var(--text-dark); 
        border-color: var(--border-color); backdrop-filter: blur(10px);
        background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23334155' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3e%3cpolyline points='6 9 12 15 18 9'%3e%3c/polyline%3e%3c/svg%3e");
    }

    /* PREMIUM THEME DEFINITIONS */
    :root, [data-theme='default'] {
        --primary-color: #6366f1; --secondary-color: #8b5cf6; --accent-color: #ec4899;
        --header-bg: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        --bg-main: #fafbfc; --bg-card: #ffffff; --bg-secondary: #f8fafc;
        --text-dark: #0f172a; --text-light: #ffffff; --text-muted: #64748b;
        --border-color: #e2e8f0; --shadow-color: rgba(99, 102, 241, 0.15);
        --success-bg: #dcfce7; --success-text: #166534; --success-border: #86efac;
        --error-bg: #fee2e2; --error-text: #dc2626; --error-border: #fca5a5;
        --gradient-accent: linear-gradient(135deg, rgba(99, 102, 241, 0.08), rgba(139, 92, 246, 0.08));
    }
    [data-theme='dark'] {
        --primary-color: #818cf8; --secondary-color: #a78bfa; --accent-color: #f472b6;
        --header-bg: linear-gradient(135deg, #0f0f23 0%, #1a1a2e 50%, #16213e 100%);
        --bg-main: #0a0a0f; --bg-card: #1a1a2e; --bg-secondary: #16213e;
        --text-dark: #f8fafc; --text-light: #ffffff; --text-muted: #94a3b8;
        --border-color: #374151; --shadow-color: rgba(129, 140, 248, 0.2);
        --gradient-accent: linear-gradient(135deg, rgba(129, 140, 248, 0.1), rgba(167, 139, 250, 0.1));
    }
    [data-theme='ocean'] {
        --primary-color: #0ea5e9; --secondary-color: #06b6d4; --accent-color: #3b82f6;
        --header-bg: linear-gradient(135deg, #0c4a6e 0%, #075985 50%, #0369a1 100%);
        --bg-main: #f0f9ff; --bg-card: #ffffff; --bg-secondary: #e0f2fe;
        --text-dark: #0c4a6e; --text-light: #ffffff; --text-muted: #475569;
        --border-color: #bae6fd; --shadow-color: rgba(14, 165, 233, 0.15);
        --gradient-accent: linear-gradient(135deg, rgba(14, 165, 233, 0.08), rgba(6, 182, 212, 0.08));
    }
    [data-theme='forest'] {
        --primary-color: #22c55e; --secondary-color: #16a34a; --accent-color: #84cc16;
        --header-bg: linear-gradient(135deg, #14532d 0%, #166534 50%, #15803d 100%);
        --bg-main: #f0fdf4; --bg-card: #ffffff; --bg-secondary: #ecfdf5;
        --text-dark: #14532d; --text-light: #ffffff; --text-muted: #4d7c0f;
        --border-color: #bbf7d0; --shadow-color: rgba(34, 197, 94, 0.15);
        --gradient-accent: linear-gradient(135deg, rgba(34, 197, 94, 0.08), rgba(22, 163, 74, 0.08));
    }
    [data-theme='sunset'] {
        --primary-color: #f97316; --secondary-color: #ea580c; --accent-color: #dc2626;
        --header-bg: linear-gradient(135deg, #9a3412 0%, #c2410c 50%, #dc2626 100%);
        --bg-main: #fffbeb; --bg-card: #ffffff; --bg-secondary: #fef3c7;
        --text-dark: #9a3412; --text-light: #ffffff; --text-muted: #a16207;
        --border-color: #fed7aa; --shadow-color: rgba(249, 115, 22, 0.15);
        --gradient-accent: linear-gradient(135deg, rgba(249, 115, 22, 0.08), rgba(234, 88, 12, 0.08));
    }
    [data-theme='rose'] {
        --primary-color: #f43f5e; --secondary-color: #ec4899;
        --header-bg: linear-gradient(135deg, #881337 0%, #9d174d 50%, #be185d 100%);
        --bg-main: #fff1f2; --bg-card: #ffffff; --text-dark: #881337; --text-muted: #9f1239;
        --border-color: #fbb6ce; --shadow-color: rgba(244, 63, 94, 0.15);
        --gradient-accent: linear-gradient(135deg, rgba(244, 63, 94, 0.1), rgba(236, 72, 153, 0.1));
    }
    [data-theme='slate'] {
        --primary-color: #64748b; --secondary-color: #475569;
        --header-bg: linear-gradient(135deg, #1e293b 0%, #334155 50%, #475569 100%);
        --bg-main: #f8fafc; --bg-card: #ffffff; --text-dark: #0f172a; --text-muted: #334155;
        --border-color: #cbd5e1; --shadow-color: rgba(100, 116, 139, 0.15);
        --gradient-accent: linear-gradient(135deg, rgba(100, 116, 139, 0.1), rgba(71, 85, 105, 0.1));
    }
    [data-theme='nebula'] {
        --primary-color: #8b5cf6; --secondary-color: #d946ef;
        --header-bg: linear-gradient(135deg, #581c87 0%, #7c3aed 50%, #a855f7 100%);
        --bg-main: #faf5ff; --bg-card: #ffffff; --text-dark: #581c87; --text-muted: #7c2d12;
        --border-color: #e9d5ff; --shadow-color: rgba(139, 92, 246, 0.15);
        --gradient-accent: linear-gradient(135deg, rgba(139, 92, 246, 0.1), rgba(217, 70, 239, 0.1));
    }
    [data-theme='lime'] {
        --primary-color: #84cc16; --secondary-color: #a3e635;
        --header-bg: linear-gradient(135deg, #365314 0%, #4d7c0f 50%, #65a30d 100%);
        --bg-main: #f7fee7; --bg-card: #ffffff; --text-dark: #365314; --text-muted: #4d7c0f;
        --border-color: #d9f99d; --shadow-color: rgba(132, 204, 22, 0.15);
        --gradient-accent: linear-gradient(135deg, rgba(132, 204, 22, 0.1), rgba(163, 230, 53, 0.1));
    }
    [data-theme='copper'] {
        --primary-color: #d97706; --secondary-color: #b45309;
        --header-bg: linear-gradient(135deg, #78350f 0%, #92400e 50%, #a16207 100%);
        --bg-main: #fefce8; --bg-card: #ffffff; --text-dark: #78350f; --text-muted: #92400e;
        --border-color: #fde68a; --shadow-color: rgba(217, 119, 6, 0.15);
        --gradient-accent: linear-gradient(135deg, rgba(217, 119, 6, 0.1), rgba(180, 83, 9, 0.1));
    }

    /* BASE STYLES */
    html { 
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif; 
        scroll-behavior: smooth; line-height: 1.6;
    }
    body {
        background: var(--bg-main); min-height: 100vh; color: var(--text-dark);
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        display: flex; align-items: center; justify-content: center; padding: 1.5rem;
        background-image: 
            radial-gradient(circle at 25% 25%, var(--primary-color)08 0%, transparent 50%),
            radial-gradient(circle at 75% 75%, var(--secondary-color)06 0%, transparent 50%);
    }

    /* CONTAINERS */
    .container, .card {
        background: var(--bg-card); border-radius: 32px; width: 100%;
        box-shadow: 0 32px 64px -12px var(--shadow-color), 0 0 0 1px var(--border-color);
        border: 1px solid var(--border-color); overflow: hidden; position: relative;
        backdrop-filter: blur(20px); transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .container:hover, .card:hover {
        transform: translateY(-8px) scale(1.01);
        box-shadow: 0 48px 80px -12px var(--shadow-color), 0 0 0 1px var(--primary-color);
    }
    .container { max-width: 1000px; }
    .card { max-width: 500px; text-align: center; padding: 3.5rem; position: relative; }

    /* TYPOGRAPHY */
    h1, h2, h3, h4 { 
        background: linear-gradient(135deg, var(--primary-color), var(--secondary-color), var(--accent-color));
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        background-clip: text; font-weight: 800; line-height: 1.2; letter-spacing: -0.025em;
    }
    h1 { font-size: clamp(2rem, 4vw, 3rem); margin-bottom: 0.75rem; }
    h2 { font-size: clamp(1.5rem, 3vw, 2.25rem); margin-bottom: 1rem; }
    h3 { font-size: clamp(1.25rem, 2.5vw, 1.75rem); margin-bottom: 0.75rem; }
    p { color: var(--text-muted); line-height: 1.7; font-size: 1.1rem; }
    .tagline { font-size: 1.2rem; opacity: 0.9; margin-top: 0.75rem; font-weight: 500; }

    /* HEADER */
    .header {
        background: var(--header-bg); color: var(--text-light); padding: 4rem 3rem;
        text-align: center; position: relative; border-radius: 32px 32px 0 0;
        overflow: hidden;
    }
    .header::before {
        content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0;
        background: linear-gradient(135deg, transparent 30%, rgba(255,255,255,0.1) 50%, transparent 70%);
        animation: shimmer 4s ease-in-out infinite;
    }
    @keyframes shimmer {
        0%, 100% { transform: translateX(-100%) rotate(-45deg); }
        50% { transform: translateX(100%) rotate(-45deg); }
    }
    .main { padding: 3.5rem; }

    /* FORMS */
    label { 
        font-weight: 600; font-size: 0.95rem; display: block; margin-bottom: 0.75rem; 
        text-align: left; color: var(--text-dark); letter-spacing: 0.025em;
    }
    input, textarea, select {
        width: 100%; padding: 1.25rem 1.5rem; border: 2px solid var(--border-color);
        background: var(--bg-secondary); color: var(--text-dark); border-radius: 20px;
        font-size: 1.05rem; transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        font-family: inherit; font-weight: 500; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
    }
    input:focus, textarea:focus, select:focus { 
        outline: none; border-color: var(--primary-color); background: var(--bg-card);
        box-shadow: 0 0 0 4px var(--shadow-color), 0 12px 32px -8px var(--shadow-color);
        transform: translateY(-2px) scale(1.01);
    }
    input:hover, textarea:hover, select:hover {
        border-color: var(--secondary-color); transform: translateY(-1px);
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.08);
    }

    /* BUTTONS */
    .btn {
        padding: 1.25rem 2.5rem; border: none; border-radius: 50px; font-weight: 700;
        cursor: pointer; text-decoration: none; display: inline-flex;
        align-items: center; justify-content: center; gap: 0.75rem; 
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        position: relative; overflow: hidden; font-family: inherit; font-size: 1rem;
        box-shadow: 0 8px 25px -8px var(--shadow-color); letter-spacing: 0.025em;
    }
    .btn::before {
        content: ''; position: absolute; top: 0; left: -100%; width: 100%; height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);
        transition: left 0.6s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .btn:hover::before { left: 100%; }
    .btn:hover { 
        transform: translateY(-4px) scale(1.05); 
        box-shadow: 0 20px 40px -8px var(--shadow-color);
    }
    .btn:active { transform: translateY(-2px) scale(1.02); }
    .btn-primary { 
        background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
        color: var(--text-light);
    }
    .btn-secondary { 
        background: var(--bg-secondary); color: var(--text-dark); 
        border: 2px solid var(--border-color); backdrop-filter: blur(10px);
    }

    /* UPLOAD AREA */
    .upload-area {
        border: 3px dashed var(--border-color); border-radius: 32px; padding: 4rem;
        cursor: pointer; transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        background: var(--gradient-accent); position: relative; overflow: hidden;
        backdrop-filter: blur(10px);
    }
    .upload-area::before {
        content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0;
        background: radial-gradient(circle at center, rgba(255,255,255,0.1) 0%, transparent 70%);
        opacity: 0; transition: opacity 0.4s ease;
    }
    .upload-area:hover::before { opacity: 1; }
    .upload-area:hover {
        border-color: var(--primary-color); transform: translateY(-6px) scale(1.02);
        box-shadow: 0 25px 50px -12px var(--shadow-color);
        background: linear-gradient(135deg, var(--gradient-accent), rgba(255,255,255,0.1));
    }
    .upload-area.dragover {
        border-color: var(--accent-color); 
        background: linear-gradient(135deg, var(--primary-color)15, var(--secondary-color)15);
        transform: scale(1.05);
    }

    /* ALERTS */
    .alert { 
        padding: 1.5rem 2rem; margin-bottom: 2rem; border-radius: 20px; 
        font-weight: 600; border: 2px solid transparent; position: relative;
        animation: slideIn 0.6s cubic-bezier(0.4, 0, 0.2, 1);
        backdrop-filter: blur(10px); font-size: 1.05rem;
    }
    @keyframes slideIn {
        from { opacity: 0; transform: translateY(-20px) scale(0.95); }
        to { opacity: 1; transform: translateY(0) scale(1); }
    }
    .alert-error { 
        background: linear-gradient(135deg, var(--error-bg), rgba(220, 38, 38, 0.1));
        color: var(--error-text); border-color: var(--error-border);
        box-shadow: 0 8px 32px -8px rgba(220, 38, 38, 0.3);
    }
    .alert-success { 
        background: linear-gradient(135deg, var(--success-bg), rgba(34, 197, 94, 0.1));
        color: var(--success-text); border-color: var(--success-border);
        box-shadow: 0 8px 32px -8px rgba(34, 197, 94, 0.3);
    }

    /* FILE INFO */
    .file-info {
        border-left: 6px solid var(--primary-color); padding: 2rem 2.5rem;
        background: var(--gradient-accent); border-radius: 0 24px 24px 0;
        margin: 2rem 0; animation: fadeIn 0.8s cubic-bezier(0.4, 0, 0.2, 1);
        backdrop-filter: blur(10px); box-shadow: 0 12px 32px -8px var(--shadow-color);
    }
    @keyframes fadeIn {
        from { opacity: 0; transform: translateX(-30px) scale(0.95); }
        to { opacity: 1; transform: translateX(0) scale(1); }
    }

    /* NAVIGATION */
    .nav-links {
        display: flex; flex-wrap: wrap; gap: 1.25rem; justify-content: center; 
        margin-top: 2rem;
    }
    .nav-links .btn {
        padding: 1rem 2rem; font-size: 0.95rem; font-weight: 600;
        background: rgba(255, 255, 255, 0.15); backdrop-filter: blur(20px);
        border: 2px solid rgba(255, 255, 255, 0.25); color: var(--text-light);
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
    }
    .nav-links .btn:hover {
        background: rgba(255, 255, 255, 0.25);
        border-color: rgba(255, 255, 255, 0.4);
        transform: translateY(-3px) scale(1.05);
    }

    /* USER INFO */
    .user-info {
        margin-top: 2.5rem; padding-top: 2rem; 
        border-top: 2px solid rgba(255, 255, 255, 0.2);
        animation: slideUp 0.8s cubic-bezier(0.4, 0, 0.2, 1) 0.3s both;
        backdrop-filter: blur(10px);
    }
    @keyframes slideUp {
        from { opacity: 0; transform: translateY(30px) scale(0.95); }
        to { opacity: 1; transform: translateY(0) scale(1); }
    }

    /* LOGIN HEADER */
    .login-header {
        margin-bottom: 3rem; animation: fadeInScale 0.8s cubic-bezier(0.4, 0, 0.2, 1);
    }
    @keyframes fadeInScale {
        from { opacity: 0; transform: scale(0.85) translateY(20px); }
        to { opacity: 1; transform: scale(1) translateY(0); }
    }

    /* FORM GROUPS */
    .form-group {
        margin-bottom: 2rem; text-align: left;
        animation: slideInLeft 0.6s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .form-group:nth-child(2) { animation-delay: 0.1s; }
    .form-group:nth-child(3) { animation-delay: 0.2s; }
    @keyframes slideInLeft {
        from { opacity: 0; transform: translateX(-30px) scale(0.95); }
        to { opacity: 1; transform: translateX(0) scale(1); }
    }

    /* GRID LAYOUT */
    .grid-form {
        display: grid; 
        grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); 
        gap: 2rem;
    }

    /* ORDER CARDS */
    .order-card {
        background: var(--gradient-accent); border-radius: 24px; padding: 2rem; 
        margin-bottom: 1.5rem; border: 2px solid var(--border-color); 
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        backdrop-filter: blur(10px); position: relative; overflow: hidden;
        box-shadow: 0 12px 32px -8px var(--shadow-color);
    }
    .order-card:hover {
        transform: translateY(-8px) scale(1.02);
        box-shadow: 0 25px 50px -12px var(--shadow-color);
        border-color: var(--primary-color);
    }

    /* RESPONSIVE DESIGN */
    @media (max-width: 768px) {
        body { padding: 1rem; }
        .container, .card { border-radius: 24px; }
        .header { padding: 3rem 2rem; }
        .main { padding: 2.5rem; }
        h1 { font-size: 2.5rem; }
        h2 { font-size: 2rem; }
        .grid-form { grid-template-columns: 1fr; }
        .nav-links { flex-direction: column; align-items: center; }
        .theme-selector { position: static; margin-bottom: 1.5rem; }
        .header-controls { position: static; justify-content: center; margin-bottom: 2rem; }
        .upload-area { padding: 3rem 2rem; }
        .btn { padding: 1rem 2rem; }
    }

    /* ANIMATIONS */
    .container, .card { 
        animation: containerFadeIn 0.8s cubic-bezier(0.4, 0, 0.2, 1); 
    }
    @keyframes containerFadeIn {
        from { opacity: 0; transform: translateY(40px) scale(0.9); }
        to { opacity: 1; transform: translateY(0) scale(1); }
    }

    /* SCROLLBAR */
    ::-webkit-scrollbar { width: 12px; }
    ::-webkit-scrollbar-track { background: var(--bg-secondary); border-radius: 10px; }
    ::-webkit-scrollbar-thumb { 
        background: linear-gradient(135deg, var(--primary-color), var(--secondary-color)); 
        border-radius: 10px; border: 2px solid var(--bg-secondary);
    }
    ::-webkit-scrollbar-thumb:hover { 
        background: linear-gradient(135deg, var(--secondary-color), var(--accent-color));
    }
</style>
<script>
    document.addEventListener('DOMContentLoaded', () => {
        const themeSwitch = document.getElementById('theme-switch');
        const currentTheme = localStorage.getItem('theme') || 'default';

        const applyTheme = (theme) => {
            document.documentElement.setAttribute('data-theme', theme);
            document.documentElement.style.transition = 'all 0.5s cubic-bezier(0.4, 0, 0.2, 1)';
            setTimeout(() => {
                document.documentElement.style.transition = '';
            }, 500);
        };

        applyTheme(currentTheme);
        if (themeSwitch) {
            themeSwitch.value = currentTheme;
            themeSwitch.addEventListener('change', (e) => {
                const selectedTheme = e.target.value;
                localStorage.setItem('theme', selectedTheme);
                applyTheme(selectedTheme);
            });
        }

        // Enhanced file drag and drop
        document.querySelectorAll('.upload-area').forEach(area => {
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                area.addEventListener(eventName, (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                }, false);
            });

            ['dragenter', 'dragover'].forEach(eventName => {
                area.addEventListener(eventName, () => {
                    area.classList.add('dragover');
                }, false);
            });

            ['dragleave', 'drop'].forEach(eventName => {
                area.addEventListener(eventName, () => {
                    area.classList.remove('dragover');
                }, false);
            });
        });
    });
</script>
"""

THEME_SELECTOR_HTML = """
<div class="theme-selector">
    <select id="theme-switch" title="Change Theme">
        <option value="default">✨ Default</option>
        <option value="dark">🌙 Dark Slate</option>
        <option value="ocean">🌊 Ocean Blue</option>
        <option value="forest">🌲 Forest Green</option>
        <option value="sunset">🌅 Sunset Orange</option>
        <option value="rose">🌹 Rose Pink</option>
        <option value="slate">⚡ Cool Slate</option>
        <option value="nebula">🌌 Nebula Purple</option>
        <option value="lime">🍃 Lime Green</option>
        <option value="copper">🔥 Copper</option>
    </select>
</div>
"""

# Base HTML structure for reuse
def create_page_template(title, body_content, is_card=False, is_container=False):
    container_class = ""
    if is_card:
        container_class = "card"
    elif is_container:
        container_class = "container"

    return f"""
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>{title} - Sale Order System</title>
      <link rel="preconnect" href="https://fonts.googleapis.com">
      <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
      <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
      {THEME_CSS_AND_JS}
    </head>
    <body><div class="{container_class}">{body_content}</div></body>
    </html>
    """

# ---------------- Helpers ----------------
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def cleanup_old_files():
    try:
        current_time = time.time()
        cleanup_threshold = app.config['CLEANUP_INTERVAL_HOURS'] * 3600
        for folder in [app.config['UPLOAD_FOLDER'], app.config['REPORT_FOLDER']]:
            if not os.path.exists(folder):
                continue
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                if os.path.isfile(file_path) and current_time - os.path.getctime(file_path) > cleanup_threshold:
                    try: os.remove(file_path)
                    except OSError as e: app.logger.error(f"Cleanup error: {e}")
    except Exception as e:
        app.logger.error(f"Error during cleanup: {e}")

def format_file_size(size_bytes):
    if size_bytes == 0: return "0 Bytes"
    k = 1024; sizes = ["Bytes","KB","MB","GB"]
    i = int(math.floor(math.log(size_bytes) / math.log(k)))
    return f"{round(size_bytes / pow(k, i), 2)} {sizes[i]}"

def get_db_connection():
    conn = sqlite3.connect(app.config['DATABASE_FILE'])
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS sale_orders (id INTEGER PRIMARY KEY,username TEXT,dealer_name TEXT,city TEXT,order_id TEXT,report_name TEXT,generated_at TEXT)")
        c.execute("CREATE TABLE IF NOT EXISTS order_id_views (id INTEGER PRIMARY KEY,username TEXT,order_id TEXT,viewed_at TEXT,ip TEXT,user_agent TEXT)")
        c.execute("CREATE TABLE IF NOT EXISTS issued_order_ids (id INTEGER PRIMARY KEY,order_id TEXT UNIQUE,given_to_name TEXT,dealer_name TEXT,city TEXT,given_by_user TEXT,given_at TEXT)")
        conn.commit()
        conn.close()
    except Exception as e:
        app.logger.error(f"DB init error: {e}")

def get_latest_order_id_global():
    try:
        conn = get_db_connection()
        row = conn.execute("SELECT order_id FROM (SELECT order_id, generated_at as ts FROM sale_orders UNION ALL SELECT order_id, given_at as ts FROM issued_order_ids) ORDER BY datetime(ts) DESC LIMIT 1").fetchone()
        conn.close()
        return row['order_id'] if row else None
    except Exception as e:
        app.logger.error(f"get_latest_order_id_global error: {e}")
        return None

def get_next_suggested_order_id():
    last_id = get_latest_order_id_global()
    if not last_id:
        return datetime.now().strftime("%m-%y-") + "00001"
    try:
        parts = last_id.split('-')
        if len(parts) == 3:
            month_year, num_part = f"{parts[0]}-{parts[1]}", parts[2]
            current_month_year = datetime.now().strftime("%m-%y")
            next_num = int(num_part) + 1 if month_year == current_month_year else 1
            return f"{current_month_year}-{next_num:05d}"
    except (ValueError, IndexError):
        pass
    return datetime.now().strftime("%m-%y-") + "00001"

# ---------------- Authentication ----------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        uname = request.form.get("username", "").strip()
        pwd = request.form.get("password", "")

        if not uname or not pwd:
            flash("Please enter both username and password.", 'error')
        elif uname in USERS and bcrypt.check_password_hash(USERS[uname], pwd):
            if uname in ACTIVE_SESSIONS:
                flash("This account is already logged in on another system.", "error")
                return redirect(url_for("login"))
            
            # Clear any existing session data
            session.clear()
            
            # Set new session
            session['user'] = uname
            session.permanent = True  # Make session permanent with timeout
            ACTIVE_SESSIONS[uname] = True
            
            app.logger.info(f"User {uname} logged in successfully")
            return redirect(url_for("home"))
        else:
            flash("Invalid credentials. Please try again.", 'error')

    messages = get_flashed_messages(with_categories=True)
    message_html = "".join(f'<div class="alert alert-{"error" if cat == "error" else "success"}">{msg}</div>' for cat, msg in messages)
    
    body = f"""
        {THEME_SELECTOR_HTML}
        <div class="login-header">
            <h2>📊 Sale Order System</h2>
            <h5>Designed & Developed by Rajesh Jadoun</h5>
            <p>Please login to continue</p>
        </div>
        {message_html}
        <form method="POST" id="loginForm">
            <div class="form-group">
                <label for="username">👤 Username</label>
                <input type="text" id="username" name="username" required autocomplete="username">
            </div>
            <div class="form-group">
                <label for="password">🔒 Password</label>
                <input type="password" id="password" name="password" required autocomplete="current-password">
            </div>
            <button type="submit" class="btn btn-primary" id="loginButton" style="width: 100%;">
                ✨ Login
            </button>
        </form>
        <script>
            document.getElementById('loginForm').addEventListener('submit', function(e) {{
                const button = document.getElementById('loginButton');
                button.disabled = true;
                button.innerHTML = '⏳ Logging in...';
                setTimeout(() => {{
                    button.disabled = false;
                    button.innerHTML = '✨ Login';
                }}, 2000);
            }});
        </script>
    """
    return create_page_template("Login", body, is_card=True)

@app.route("/logout")
def logout():
    uname = session.get("user")
    if uname:
        ACTIVE_SESSIONS.pop(uname, None)
        app.logger.info(f"User {uname} logged out")
    
    session.clear()
    flash("You have been logged out successfully.", 'success')
    return redirect(url_for("login"))

# ---------------- Decorators (STRENGTHENED) ----------------
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # Check if user is in session
        if "user" not in session:
            flash("Please login to access this page.", "error")
            return redirect(url_for("login"))
        
        # Check if user is in active sessions
        username = session.get("user")
        if username not in ACTIVE_SESSIONS:
            session.clear()  # Clear invalid session
            flash("Your session has expired. Please login again.", "error")
            return redirect(url_for("login"))
        
        # Verify user still exists in USERS
        if username not in USERS:
            session.clear()
            ACTIVE_SESSIONS.pop(username, None)
            flash("User account not found. Please contact administrator.", "error")
            return redirect(url_for("login"))
            
        return f(*args, **kwargs)
    return decorated_function

# ---------------- Routes ----------------
@app.route('/', methods=['GET', 'POST'])
@login_required
def home():
    if request.method == 'POST':
        try:
            cleanup_old_files()
            if 'file' not in request.files or request.files['file'].filename == '':
                flash("No file was selected. Please choose a file to upload.", 'error')
                return redirect(url_for('home'))
            file = request.files['file']
            if not allowed_file(file.filename):
                flash("Invalid file type. Please upload Excel files only (.xls or .xlsx).", 'error')
                return redirect(url_for('home'))
            
            filename = secure_filename(file.filename)
            unique_filename = f"{uuid.uuid4().hex}_{filename}"
            temp_input_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(temp_input_path)
            
            session['temp_file_path'] = temp_input_path
            session['uploaded_filename'] = filename
            session['uploaded_filesize'] = format_file_size(os.path.getsize(temp_input_path))
            session['upload_time'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            return redirect(url_for('order_form'))
        except Exception as e:
            current_app.logger.error(f"Upload error: {e}")
            flash("There was an error processing your file.", "error")
            return redirect(url_for('home'))

    body = f"""
    <div class="header">
        <div class="header-controls">{THEME_SELECTOR_HTML}</div>
        <h1><b>📊 Sale Order System</b></h1>
        <p class="tagline"><b>Accurate Orders, Better Business ✨</b></p>
       <h5><strong>Designed & Developed by Rajesh Jadoun ✨</strong></h5>
        <div class="user-info">
            <span>Welcome back, <strong>{session['user']}</strong>! 👋</span>
            <div class="nav-links">
                <a href="/" class="btn btn-secondary">🏠 Home</a>
                <a href="/orders" class="btn btn-secondary">📋 My Orders</a>
                <a href="/last-id" class="btn btn-secondary">🔢 Last Order ID</a>
                <a href="/issue-order-id" class="btn btn-secondary">🎯 Give Order ID</a>
                <a href="/logout" class="btn btn-secondary">🚪 Logout</a>
            </div>
        </div>
    </div>
    <div class="main">
        <form method="POST" enctype="multipart/form-data" id="uploadForm">
            <div class="upload-area" id="uploadArea">
                <div style="font-size: 4rem; margin-bottom: 1rem;">📄</div>
                <div style="font-size: 1.5rem; font-weight: 700; margin-bottom: 0.5rem; color: var(--primary-color);">
                    Drag & Drop or Click to Upload
                </div>
                <div style="color: var(--text-muted); font-size: 1rem;">
                    Supports .xls and .xlsx files • Maximum 16MB
                </div>
                <input type="file" name="file" id="fileInput" class="file-input" style="display: none;" accept=".xls,.xlsx" required>
            </div>
            <div class="file-info" id="fileInfo" style="display: none; text-align: left;">
                <div style="display: flex; align-items: center; gap: 1rem; flex-wrap: wrap;">
                    <div style="font-size: 2rem;">📊</div>
                    <div style="flex: 1; min-width: 200px;">
                        <div style="font-weight: 700; color: var(--primary-color); margin-bottom: 0.25rem;">
                            <span id="fileName"></span>
                        </div>
                        <div style="color: var(--text-muted); font-size: 0.9rem;">
                            Size: <span id="fileSize"></span> • Ready to process
                        </div>
                    </div>
                    <div style="color: var(--success-text); font-size: 1.5rem;">✅</div>
                </div>
            </div>
            <div style="text-align: center; margin-top: 2rem;">
                <button type="submit" class="upload-button btn btn-primary" id="uploadButton" style="display: none;">
                    🚀 Upload & Continue
                </button>
            </div>
        </form>
    </div>
    
    <script>
        // Page loading optimization - remove heavy animations on load
        document.addEventListener('DOMContentLoaded', function() {{{{
            // Fast file handling without heavy animations
            const uploadArea=document.getElementById('uploadArea'),
                  fileInput=document.getElementById('fileInput'),
                  fileInfo=document.getElementById('fileInfo'),
                  fileName=document.getElementById('fileName'),
                  fileSize=document.getElementById('fileSize'),
                  uploadButton=document.getElementById('uploadButton');
            
            function formatFileSize(bytes) {{{{
                if(bytes === 0) return '0 Bytes';
                const k = 1024, sizes = ['Bytes','KB','MB','GB'];
                const i = Math.floor(Math.log(bytes) / Math.log(k));
                return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
            }}}}
            
            function handleFile(file) {{{{
                const allowedTypes = ['application/vnd.ms-excel','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
                const maxSize = 16 * 1024 * 1024;
                
                if (!allowedTypes.includes(file.type) && !file.name.match(/\\.(xls|xlsx)$/i)) {{{{
                    alert('📋 Please select a valid Excel file (.xls or .xlsx)');
                    return false;
                }}}}
                if (file.size > maxSize) {{{{
                    alert('⚠️ File size must be less than 16MB');
                    return false;
                }}}}
                
                fileName.textContent = file.name;
                fileSize.textContent = formatFileSize(file.size);
                fileInfo.style.display = 'block';
                uploadButton.style.display = 'inline-flex';
                uploadArea.style.borderColor = 'var(--primary-color)';
                return true;
            }}}}
            
            uploadArea.onclick = () => fileInput.click();
            fileInput.onchange = e => {{{{
                const file = e.target.files[0];
                if(file) handleFile(file);
            }}}};
            
            uploadArea.ondragover = e => {{{{
                e.preventDefault();
                uploadArea.style.borderColor = 'var(--primary-color)';
            }}}};
            uploadArea.ondragleave = e => {{{{
                e.preventDefault();
                uploadArea.style.borderColor = 'var(--border-color)';
            }}}};
            uploadArea.ondrop = e => {{{{
                e.preventDefault();
                uploadArea.style.borderColor = 'var(--border-color)';
                const file = e.dataTransfer.files[0];
                if(file) {{{{
                    fileInput.files = e.dataTransfer.files;
                    handleFile(file);
                }}}}
            }}}};
            
            document.getElementById('uploadForm').onsubmit = function(e) {{{{
                const button = document.getElementById('uploadButton');
                button.disabled = true;
                button.innerHTML = '⏳ Processing...';
            }}}};
        }}}});
    </script>
    """
    return create_page_template("Dashboard", body, is_container=True)

@app.route('/form', methods=['GET','POST'])
@login_required
def order_form():
    if 'temp_file_path' not in session or not os.path.exists(session['temp_file_path']):
        flash("Please upload an Excel file first.", "error")
        return redirect(url_for('home'))

    if request.method == 'POST':
        try:
            dealer_name = request.form.get('dealer_name','').strip()
            city = request.form.get('city','').strip()
            order_date_raw = request.form.get('order_date','')
            order_date = datetime.strptime(order_date_raw, '%Y-%m-%d').strftime('%d-%m-%Y') if order_date_raw else ""
            freight_condition = request.form.get('freight_condition','').strip()

            df, _, weight_map, hdmr_map, mdf_map, ply_map, pvc_map, wpc_map = prepare_data(session['temp_file_path'])
            
            safe_dealer_name = re.sub(r'[^a-zA-Z0-9_.-]+', '_', dealer_name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_name = f"{safe_dealer_name}_{timestamp}_SALE_ORDER.xlsx"
            output_path = os.path.join(app.config['REPORT_FOLDER'], report_name)

            write_report(df, output_path, weight_map, hdmr_map, mdf_map, ply_map, pvc_map, wpc_map, 
                         session['user'], dealer_name, city, order_date, freight_condition)

            os.remove(session['temp_file_path'])
            session.pop('temp_file_path', None)
            
            return redirect(url_for('download_report', report_name=report_name))
        except Exception as e:
            current_app.logger.error(f"Generate error: {e}")
            flash("Failed to generate the report. An unexpected error occurred.", "error")
    
    body = f"""
    <div class="header">
        <div class="header-controls">{THEME_SELECTOR_HTML}</div>
        <h1>📝 Order Details</h1>
        <p>Fill in the information to generate your professional sale order</p>
    </div>
    <div class="main">
        <div class="file-info" style="margin-bottom: 2.5rem;">
            <div style="display: flex; align-items: center; gap: 1rem;">
                <div style="font-size: 2rem;">📊</div>
                <div>
                    <div style="font-weight: 700; color: var(--primary-color);">
                        {session.get('uploaded_filename','N/A')}
                    </div>
                    <div style="color: var(--text-muted); font-size: 0.9rem;">
                        Size: {session.get('uploaded_filesize','N/A')} • Uploaded: {session.get('upload_time','N/A')}
                    </div>
                </div>
            </div>
        </div>
        
        <form method="POST" id="orderForm">
            <div class="grid-form">
                <div class="form-group">
                    <label for="dealer_name">🏢 Dealer Name *</label>
                    <input id="dealer_name" name="dealer_name" required placeholder="Enter dealer name">
                </div>
                <div class="form-group">
                    <label for="city">🌍 City *</label>
                    <input id="city" name="city" required placeholder="Enter city name">
                </div>
                <div class="form-group">
                    <label for="order_date">📅 Order Date *</label>
                    <input type="date" id="order_date" name="order_date" required>
                </div>
                <div class="form-group">
                    <label for="freight_condition">🚚 Freight Condition</label>
                    <input id="freight_condition" name="freight_condition" placeholder="e.g., FOB, CIF, Ex-Works">
                </div>
            </div>
            
            <div class="form-group">
                <label for="additional_notes">📋 Additional Notes</label>
                <textarea id="additional_notes" name="additional_notes" rows="3" 
                    placeholder="Any special instructions or notes for this order..."></textarea>
            </div>
            
            <div style="display: flex; gap: 1.5rem; justify-content: center; margin-top: 2.5rem; flex-wrap: wrap;">
                <a class="btn btn-secondary" href="/">
                    ← Back to Home
                </a>
                <button class="btn btn-primary" type="submit" id="generateButton">
                    🚀 Generate Report
                </button>
            </div>
        </form>
    </div>
    
    <script>
        // Set today's date as default
        document.getElementById('order_date').valueAsDate = new Date();
        
        // Form submission handling
        document.getElementById('orderForm').addEventListener('submit', function(e) {{{{
            const button = document.getElementById('generateButton');
            const dealerName = document.getElementById('dealer_name').value.trim();
            const city = document.getElementById('city').value.trim();
            
            if (!dealerName || !city) {{{{
                e.preventDefault();
                alert('⚠️ Please fill in all required fields');
                return;
            }}}}
            
            button.disabled = true;
            button.innerHTML = '⏳ Generating Report...';
            
            // Add loading animation
            setTimeout(() => {{{{
                if (!button.disabled) return;
                button.innerHTML = '📊 Processing Data...';
            }}}}, 1000);
        }}}});
        
        // Enhanced form validation with real-time feedback
        const inputs = document.querySelectorAll('input[required]');
        inputs.forEach(input => {{{{
            input.addEventListener('blur', function() {{{{
                if (this.value.trim()) {{{{
                    this.style.borderColor = 'var(--success-text)';
                }}}} else {{{{
                    this.style.borderColor = 'var(--error-text)';
                }}}}
            }}}});
            input.addEventListener('focus', function() {{{{
                this.style.borderColor = 'var(--primary-color)';
            }}}});
        }}}});
    </script>
    """
    return create_page_template("Order Details", body, is_container=True)

@app.route('/download/<path:report_name>')
@login_required
def download_report(report_name):
    path = os.path.join(app.config['REPORT_FOLDER'], report_name)
    if not os.path.exists(path):
        flash("Report not found.", "error")
        return redirect(url_for('home'))
    return send_file(path, as_attachment=True)

@app.route('/orders')
@login_required
def orders():
    try:
        conn = get_db_connection()
        orders = conn.execute(
            "SELECT * FROM sale_orders WHERE username = ? ORDER BY datetime(generated_at) DESC LIMIT 50", 
            (session['user'],)
        ).fetchall()
        conn.close()
    except Exception as e:
        current_app.logger.error(f"Orders fetch error: {e}")
        orders = []

    orders_html = ""
    if orders:
        for order in orders:
            orders_html += f"""
            <div class="order-card" style="background: var(--gradient-accent); border-radius: 16px; padding: 1.5rem; margin-bottom: 1rem; border: 1px solid var(--border-color); transition: all 0.3s ease;">
                <div style="display: flex; justify-content: between; align-items: center; flex-wrap: wrap; gap: 1rem;">
                    <div style="flex: 1; min-width: 200px;">
                        <h3 style="color: var(--primary-color); margin-bottom: 0.5rem;">📋 {order['dealer_name']}</h3>
                        <p style="margin-bottom: 0.25rem;"><strong>🌍 City:</strong> {order['city']}</p>
                        <p style="margin-bottom: 0.25rem;"><strong>🔢 Order ID:</strong> {order['order_id']}</p>
                        <p style="color: var(--text-muted); font-size: 0.9rem;">📅 Generated: {order['generated_at']}</p>
                    </div>
                    <div style="display: flex; gap: 0.75rem; flex-wrap: wrap;">
                        <a href="/download/{order['report_name']}" class="btn btn-primary" style="padding: 0.5rem 1rem; font-size: 0.9rem;">
                            📥 Download
                        </a>
                    </div>
                </div>
            </div>
            """
    else:
        orders_html = """
        <div style="text-align: center; padding: 3rem; color: var(--text-muted);">
            <div style="font-size: 4rem; margin-bottom: 1rem; opacity: 0.5;">📋</div>
            <h3 style="margin-bottom: 0.5rem;">No Orders Yet</h3>
            <p>Your generated orders will appear here</p>
            <a href="/" class="btn btn-primary" style="margin-top: 1.5rem;">🚀 Create First Order</a>
        </div>
        """

    body = f"""
    <div class="header">
        <div class="header-controls">{THEME_SELECTOR_HTML}</div>
        <h1>📋 My Orders</h1>
        <p>View and download your generated sale orders</p>
        <div class="nav-links">
            <a href="/" class="btn btn-secondary">🏠 Home</a>
            <a href="/last-id" class="btn btn-secondary">🔢 Last Order ID</a>
            <a href="/logout" class="btn btn-secondary">🚪 Logout</a>
        </div>
    </div>
    <div class="main">
        {orders_html}
    </div>
    <style>
        .order-card:hover {{{{
            transform: translateY(-2px);
            box-shadow: 0 8px 25px -8px var(--shadow-color);
        }}}}
    </style>
    """
    return create_page_template("My Orders", body, is_container=True)

@app.route('/last-id')
@login_required
def last_id():
    try:
        latest_id = get_latest_order_id_global()
        suggested_id = get_next_suggested_order_id()
        
        conn = get_db_connection()
        recent_orders = conn.execute(
            "SELECT order_id, dealer_name, city, generated_at FROM sale_orders ORDER BY datetime(generated_at) DESC LIMIT 10"
        ).fetchall()
        conn.close()
    except Exception as e:
        current_app.logger.error(f"Last ID fetch error: {e}")
        latest_id = None
        suggested_id = get_next_suggested_order_id()
        recent_orders = []

    recent_html = ""
    if recent_orders:
        for order in recent_orders:
            recent_html += f"""
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 1rem; border-bottom: 1px solid var(--border-color); transition: all 0.2s ease;">
                <div style="flex: 1;">
                    <div style="font-weight: 600; color: var(--primary-color); margin-bottom: 0.25rem;">
                        {order['order_id']}
                    </div>
                    <div style="font-size: 0.9rem; color: var(--text-muted);">
                        {order['dealer_name']} • {order['city']}
                    </div>
                </div>
                <div style="text-align: right; color: var(--text-muted); font-size: 0.85rem;">
                    {order['generated_at']}
                </div>
            </div>
            """
    else:
        recent_html = """
        <div style="text-align: center; padding: 2rem; color: var(--text-muted);">
            <div style="font-size: 2.5rem; margin-bottom: 0.5rem; opacity: 0.5;">🔢</div>
            <p>No recent orders found</p>
        </div>
        """

    body = f"""
    <div class="header">
        <div class="header-controls">{THEME_SELECTOR_HTML}</div>
        <h1>🔢 Order ID Status</h1>
        <p>Track your order numbering system</p>
        <div class="nav-links">
            <a href="/" class="btn btn-secondary">🏠 Home</a>
            <a href="/orders" class="btn btn-secondary">📋 My Orders</a>
            <a href="/issue-order-id" class="btn btn-secondary">🎯 Give Order ID</a>
            <a href="/logout" class="btn btn-secondary">🚪 Logout</a>
        </div>
    </div>
    <div class="main">
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 2rem; margin-bottom: 2.5rem;">
            <div style="background: var(--gradient-accent); border-radius: 20px; padding: 2rem; border: 1px solid var(--border-color); text-align: center;">
                <div style="font-size: 3rem; margin-bottom: 1rem;">🏷️</div>
                <h3 style="margin-bottom: 0.5rem; color: var(--primary-color);">Latest Order ID</h3>
                <div style="font-size: 1.5rem; font-weight: 700; color: var(--text-dark); margin-bottom: 0.5rem;">
                    {latest_id or "None yet"}
                </div>
                <p style="color: var(--text-muted); font-size: 0.9rem;">Most recently used</p>
            </div>
            
            <div style="background: var(--gradient-accent); border-radius: 20px; padding: 2rem; border: 1px solid var(--border-color); text-align: center;">
                <div style="font-size: 3rem; margin-bottom: 1rem;">🎯</div>
                <h3 style="margin-bottom: 0.5rem; color: var(--primary-color);">Next Suggested ID</h3>
                <div style="font-size: 1.5rem; font-weight: 700; color: var(--text-dark); margin-bottom: 0.5rem;">
                    {suggested_id}
                </div>
                <p style="color: var(--text-muted); font-size: 0.9rem;">Auto-generated sequence</p>
            </div>
        </div>
        
        <div style="background: var(--bg-card); border-radius: 20px; border: 1px solid var(--border-color); overflow: hidden;">
            <div style="padding: 1.5rem; border-bottom: 1px solid var(--border-color); background: var(--gradient-accent);">
                <h3 style="color: var(--primary-color); margin-bottom: 0.5rem;">📊 Recent Order IDs</h3>
                <p style="color: var(--text-muted); margin: 0;">Latest 10 generated orders</p>
            </div>
            <div style="max-height: 400px; overflow-y: auto;">
                {recent_html}
            </div>
        </div>
    </div>
    """
    return create_page_template("Order ID Status", body, is_container=True)

@app.route('/issue-order-id', methods=['GET', 'POST'])
@login_required
def issue_order_id():
    if request.method == 'POST':
        try:
            order_id = request.form.get('order_id', '').strip()
            given_to_name = request.form.get('given_to_name', '').strip()
            dealer_name = request.form.get('dealer_name', '').strip()
            city = request.form.get('city', '').strip()
            
            if not all([order_id, given_to_name]):
                flash("Please fill in all required fields.", "error")
                return redirect(url_for('issue_order_id'))
            
            conn = get_db_connection()
            # Check if order ID already exists
            existing = conn.execute("SELECT * FROM issued_order_ids WHERE order_id = ?", (order_id,)).fetchone()
            if existing:
                flash(f"Order ID {order_id} has already been issued.", "error")
                conn.close()
                return redirect(url_for('issue_order_id'))
            
            # Insert new issued order ID
            conn.execute(
                "INSERT INTO issued_order_ids (order_id, given_to_name, dealer_name, city, given_by_user, given_at) VALUES (?, ?, ?, ?, ?, ?)",
                (order_id, given_to_name, dealer_name, city, session['user'], datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            )
            conn.commit()
            conn.close()
            
            flash(f"Order ID {order_id} has been successfully issued to {given_to_name}.", "success")
            return redirect(url_for('issue_order_id_success', order_id=order_id, given_to=given_to_name))
            
        except Exception as e:
            current_app.logger.error(f"Issue order ID error: {e}")
            flash("An error occurred while issuing the order ID.", "error")

    # Get suggested next ID and recent issued IDs
    try:
        suggested_id = get_next_suggested_order_id()
        conn = get_db_connection()
        recent_issued = conn.execute(
            "SELECT * FROM issued_order_ids ORDER BY datetime(given_at) DESC LIMIT 10"
        ).fetchall()
        conn.close()
    except Exception as e:
        current_app.logger.error(f"Fetch issued IDs error: {e}")
        suggested_id = get_next_suggested_order_id()
        recent_issued = []

    messages = get_flashed_messages(with_categories=True)
    message_html = "".join(f'<div class="alert alert-{"error" if cat == "error" else "success"}">{msg}</div>' for cat, msg in messages)

    recent_html = ""
    if recent_issued:
        for issued in recent_issued:
            recent_html += f"""
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 1rem; border-bottom: 1px solid var(--border-color);">
                <div style="flex: 1;">
                    <div style="font-weight: 600; color: var(--primary-color); margin-bottom: 0.25rem;">
                        {issued['order_id']}
                    </div>
                    <div style="font-size: 0.9rem; color: var(--text-muted);">
                        Given to: {issued['given_to_name']} • {issued['dealer_name'] or 'N/A'}
                    </div>
                </div>
                <div style="text-align: right; color: var(--text-muted); font-size: 0.85rem;">
                    {issued['given_at']}<br>
                    <small>by {issued['given_by_user']}</small>
                </div>
            </div>
            """

    body = f"""
    <div class="header">
        <div class="header-controls">{THEME_SELECTOR_HTML}</div>
        <h1>🎯 Issue Order ID</h1>
        <p>Assign order IDs to team members or dealers</p>
        <div class="nav-links">
            <a href="/" class="btn btn-secondary">🏠 Home</a>
            <a href="/orders" class="btn btn-secondary">📋 My Orders</a>
            <a href="/last-id" class="btn btn-secondary">🔢 Last Order ID</a>
            <a href="/logout" class="btn btn-secondary">🚪 Logout</a>
        </div>
    </div>
    <div class="main">
        {message_html}
        
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 2rem; margin-bottom: 2rem;">
            <div style="background: var(--gradient-accent); border-radius: 20px; padding: 2rem; border: 1px solid var(--border-color);">
                <h3 style="color: var(--primary-color); margin-bottom: 1rem;">📝 Issue New Order ID</h3>
                <form method="POST" id="issueForm">
                    <div class="form-group">
                        <label for="order_id">🔢 Order ID *</label>
                        <input type="text" id="order_id" name="order_id" value="{suggested_id}" required>
                    </div>
                    <div class="form-group">
                        <label for="given_to_name">👤 Given To *</label>
                        <input type="text" id="given_to_name" name="given_to_name" placeholder="Person's name" required>
                    </div>
                    <div class="form-group">
                        <label for="dealer_name">🏢 Dealer Name</label>
                        <input type="text" id="dealer_name" name="dealer_name" placeholder="Optional dealer name">
                    </div>
                    <div class="form-group">
                        <label for="city">🌍 City</label>
                        <input type="text" id="city" name="city" placeholder="Optional city">
                    </div>
                    <button type="submit" class="btn btn-primary" style="width: 100%;">
                        🎯 Issue Order ID
                    </button>
                </form>
            </div>
            
            <div style="background: var(--bg-card); border-radius: 20px; border: 1px solid var(--border-color); overflow: hidden;">
                <div style="padding: 1.5rem; border-bottom: 1px solid var(--border-color); background: var(--gradient-accent);">
                    <h3 style="color: var(--primary-color); margin-bottom: 0.5rem;">📊 Recently Issued</h3>
                    <p style="color: var(--text-muted); margin: 0;">Last 10 issued order IDs</p>
                </div>
                <div style="max-height: 400px; overflow-y: auto;">
                    {recent_html if recent_html else '<div style="text-align: center; padding: 2rem; color: var(--text-muted);">No issued order IDs yet</div>'}
                </div>
            </div>
        </div>
    </div>
    
    <script>
        document.getElementById('issueForm').addEventListener('submit', function(e) {{{{
            const orderId = document.getElementById('order_id').value.trim();
            const givenTo = document.getElementById('given_to_name').value.trim();
            
            if (!orderId || !givenTo) {{{{
                e.preventDefault();
                alert('⚠️ Please fill in all required fields');
                return;
            }}}}
        }}}});
    </script>
    """
    return create_page_template("Issue Order ID", body, is_container=True)

@app.route('/issue-success')
@login_required
def issue_order_id_success():
    order_id = request.args.get('order_id', '')
    given_to = request.args.get('given_to', '')
    
    if not order_id or not given_to:
        return redirect(url_for('home'))

    body = f"""
    <div style="text-align: center; padding: 3rem; background: var(--gradient-accent); border-radius: 24px; border: 1px solid var(--border-color);">
        <div style="font-size: 5rem; margin-bottom: 1.5rem; animation: checkmark 0.6s ease-in-out;">✅</div>
        <h1 style="color: var(--primary-color); margin-bottom: 1rem;">Order ID Issued Successfully!</h1>
        <div style="background: var(--bg-card); padding: 2rem; border-radius: 16px; margin: 2rem 0; border: 1px solid var(--border-color);">
            <h3 style="color: var(--secondary-color); margin-bottom: 1rem;">📋 Details</h3>
            <p style="font-size: 1.2rem; font-weight: 600; color: var(--primary-color); margin-bottom: 0.5rem;">
                Order ID: <span style="background: var(--gradient-accent); padding: 0.5rem 1rem; border-radius: 8px;">{order_id}</span>
            </p>
            <p style="font-size: 1.1rem; color: var(--text-dark);">
                Given to: <strong>{given_to}</strong>
            </p>
        </div>
        
        <div style="margin: 2rem 0; padding: 1.5rem; background: var(--success-bg); border-radius: 12px; border: 1px solid var(--success-text);">
            <div style="font-size: 1.5rem; margin-bottom: 0.5rem;">⏰</div>
            <p style="color: var(--success-text); font-weight: 600;">
                Redirecting to Home in <span id="countdown" style="font-size: 1.2rem; font-weight: 700;">20</span> seconds...
            </p>
            <div style="width: 100%; height: 6px; background: rgba(0,0,0,0.1); border-radius: 3px; margin: 1rem 0; overflow: hidden;">
                <div id="progress-bar" style="height: 100%; background: var(--success-text); border-radius: 3px; width: 100%; transition: width 20s linear;"></div>
            </div>
        </div>
        
        <div style="display: flex; gap: 1rem; justify-content: center; flex-wrap: wrap;">
            <a href="/" class="btn btn-primary">🏠 Go to Home Now</a>
            <a href="/issue-order-id" class="btn btn-secondary">🎯 Issue Another ID</a>
            <button id="cancelTimer" class="btn btn-secondary">⏸️ Cancel Timer</button>
        </div>
    </div>
    
    <style>
        @keyframes checkmark {{{{
            0% {{{{ transform: scale(0) rotate(0deg); opacity: 0; }}}}
            50% {{{{ transform: scale(1.2) rotate(180deg); opacity: 0.8; }}}}
            100% {{{{ transform: scale(1) rotate(360deg); opacity: 1; }}}}
        }}}}
        
        @keyframes pulse {{{{
            0%, 100% {{{{ transform: scale(1); }}}}
            50% {{{{ transform: scale(1.05); }}}}
        }}}}
        
        #countdown {{{{ 
            animation: pulse 1s infinite; 
            color: var(--error-text); 
        }}}}
    </style>
    
    <script>
        let timeLeft = 20;
        let timerId;
        let cancelled = false;
        
        function updateCountdown() {{{{
            if (cancelled) return;
            
            const countdownEl = document.getElementById('countdown');
            const progressBar = document.getElementById('progress-bar');
            
            if (timeLeft > 0) {{{{
                countdownEl.textContent = timeLeft;
                const progressWidth = (timeLeft / 20) * 100;
                progressBar.style.width = progressWidth + '%';
                timeLeft--;
            }}}} else {{{{
                window.location.href = '/';
            }}}}
        }}}}
        
        // Start countdown
        document.addEventListener('DOMContentLoaded', function() {{{{
            // Start progress bar animation
            setTimeout(() => {{{{
                document.getElementById('progress-bar').style.width = '0%';
            }}}}, 100);
            
            // Update countdown every second
            updateCountdown();
            timerId = setInterval(updateCountdown, 1000);
        }}}});
        
        // Cancel timer button
        document.getElementById('cancelTimer').addEventListener('click', function() {{{{
            cancelled = true;
            clearInterval(timerId);
            document.getElementById('progress-bar').style.width = '100%';
            document.getElementById('countdown').textContent = '∞';
            this.style.display = 'none';
            
            // Show success message
            const successMsg = document.createElement('div');
            successMsg.innerHTML = '⏸️ <strong>Timer Cancelled!</strong> Stay on this page as long as you want.';
            successMsg.style.cssText = 'margin-top: 1rem; padding: 1rem; background: var(--gradient-accent); border-radius: 8px; color: var(--primary-color); font-weight: 600;';
            this.parentNode.appendChild(successMsg);
        }}}});
        
        // Keyboard shortcut - ESC to cancel timer
        document.addEventListener('keydown', function(e) {{{{
            if (e.key === 'Escape' && !cancelled) {{{{
                document.getElementById('cancelTimer').click();
            }}}}
        }}}});
    </script>
    """
    
    return create_page_template("Order ID Issued Successfully", body, is_card=True)

@app.route('/favicon.ico')
def favicon():
    # Return a simple SVG favicon
    svg_content = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">
        <rect width="64" height="64" rx="12" fill="#6366f1"/>
        <text x="32" y="45" font-family="Arial, sans-serif" font-size="36" font-weight="bold" text-anchor="middle" fill="white">📊</text>
    </svg>'''
    return svg_content, 200, {'Content-Type': 'image/svg+xml'}
