import streamlit as st
import pandas as pd
import json
import os
import time
import zipfile
import io
from datetime import datetime
from pathlib import Path
import sys
import subprocess
import threading
import asyncio
import logging

# Set page configuration
st.set_page_config(
    page_title="CRM Invoice Automation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #28a745;
    }
    .error-box {
        background-color: #f8d7da;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #dc3545;
    }
    .info-box {
        background-color: #d1ecf1;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #17a2b8;
    }
    .stButton > button {
        width: 100%;
        margin-bottom: 10px;
    }
    .tab-content {
        padding: 20px 0;
    }
</style>
""", unsafe_allow_html=True)

class StreamlitCRMAutomation:
    def __init__(self):
        self.config_file = "config.json"
        self.download_dir = "invoices"
        self.config = self.load_config()
        
        # Ensure directories exist
        Path(self.download_dir).mkdir(exist_ok=True)
        Path("logs").mkdir(exist_ok=True)
        
        # Session state initialization
        if 'automation_running' not in st.session_state:
            st.session_state.automation_running = False
        if 'automation_logs' not in st.session_state:
            st.session_state.automation_logs = []
        if 'stats' not in st.session_state:
            st.session_state.stats = {
                'total': 0,
                'success': 0,
                'failed': 0,
                'start_time': None,
                'end_time': None
            }
        if 'current_step' not in st.session_state:
            st.session_state.current_step = None
        if 'processed_so' not in st.session_state:
            st.session_state.processed_so = []

    def load_config(self):
        """Load existing configuration or create default"""
        default_config = {
            "crm_url": "",
            "username": "",
            "password": "",
            "excel_path": "",
            "download_path": "invoices",
            "headless": True,
            "wait_time": 5,
            "max_retries": 3,
            "log_level": "INFO"
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, "r") as f:
                    loaded_config = json.load(f)
                    # Merge with defaults
                    for key in default_config:
                        if key in loaded_config:
                            default_config[key] = loaded_config[key]
        except Exception as e:
            st.error(f"Error loading config: {e}")
            
        return default_config

    def save_config(self):
        """Save configuration to file"""
        try:
            with open(self.config_file, "w") as f:
                json.dump(self.config, f, indent=4)
            return True
        except Exception as e:
            st.error(f"Failed to save configuration: {e}")
            return False

    def add_log(self, message, level="INFO"):
        """Add log message to session state"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}"
        st.session_state.automation_logs.append(log_entry)
        
        # Keep only last 100 logs
        if len(st.session_state.automation_logs) > 100:
            st.session_state.automation_logs = st.session_state.automation_logs[-100:]

    def render_sidebar(self):
        """Render sidebar with status and quick actions"""
        with st.sidebar:
            st.title("📋 Dashboard")
            
            # Status indicator
            if st.session_state.automation_running:
                st.warning("🔄 Automation Running")
            else:
                st.success("✅ Ready")
            
            st.divider()
            
            # Quick Stats
            st.subheader("Statistics")
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total", st.session_state.stats['total'])
            with col2:
                st.metric("Success", st.session_state.stats['success'])
            
            st.divider()
            
            # Quick Actions
            st.subheader("Quick Actions")
            
            if st.button("📥 Download All Invoices", use_container_width=True):
                self.download_all_invoices()
            
            if st.button("📋 View Logs", use_container_width=True):
                st.session_state.show_logs = True
            
            if st.button("🔄 Reset Session", use_container_width=True):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
            
            st.divider()
            
            # Configuration Status
            st.subheader("Config Status")
            config_status = self.validate_config_silent()
            if config_status['valid']:
                st.success("✅ Configuration Valid")
            else:
                st.error("❌ Configuration Incomplete")
                with st.expander("Missing Items"):
                    for item in config_status['errors']:
                        st.write(f"• {item}")

    def render_main_content(self):
        """Render main content with tabs"""
        st.markdown('<h1 class="main-header">CRM Invoice Automation</h1>', unsafe_allow_html=True)
        
        # Create tabs
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "🏠 Dashboard",
            "⚙️ Configuration", 
            "📊 File Management",
            "🚀 Run Automation",
            "📋 Logs & Results"
        ])
        
        with tab1:
            self.render_dashboard()
        
        with tab2:
            self.render_configuration()
        
        with tab3:
            self.render_file_management()
        
        with tab4:
            self.render_automation()
        
        with tab5:
            self.render_logs_results()

    def render_dashboard(self):
        """Render dashboard tab"""
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📈 Overview")
            
            # Stats cards
            stats_col1, stats_col2, stats_col3 = st.columns(3)
            with stats_col1:
                st.metric("Total Orders", st.session_state.stats['total'])
            with stats_col2:
                success_rate = (st.session_state.stats['success'] / st.session_state.stats['total'] * 100) if st.session_state.stats['total'] > 0 else 0
                st.metric("Success Rate", f"{success_rate:.1f}%")
            with stats_col3:
                st.metric("Failed", st.session_state.stats['failed'])
        
        with col2:
            st.subheader("🔔 Recent Activity")
            if st.session_state.automation_logs:
                for log in st.session_state.automation_logs[-5:]:
                    st.text(log)
            else:
                st.info("No activity yet")
        
        st.divider()
        
        # Quick Start Section
        st.subheader("🚀 Quick Start")
        
        qcol1, qcol2, qcol3 = st.columns(3)
        
        with qcol1:
            if st.button("⚙️ Configure CRM", use_container_width=True):
                st.session_state.active_tab = "Configuration"
                st.rerun()
        
        with qcol2:
            if st.button("📁 Upload Excel", use_container_width=True):
                st.session_state.active_tab = "File Management"
                st.rerun()
        
        with qcol3:
            if st.button("▶️ Run Automation", use_container_width=True):
                st.session_state.active_tab = "Run Automation"
                st.rerun()

    def render_configuration(self):
        """Render configuration tab"""
        st.subheader("CRM Configuration")
        
        with st.form("config_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                crm_url = st.text_input(
                    "CRM URL *",
                    value=self.config.get("crm_url", ""),
                    help="Full URL of your CRM login page"
                )
                
                username = st.text_input(
                    "Username *",
                    value=self.config.get("username", ""),
                    help="Your CRM username"
                )
                
                password = st.text_input(
                    "Password *",
                    value=self.config.get("password", ""),
                    type="password",
                    help="Your CRM password"
                )
            
            with col2:
                headless = st.checkbox(
                    "Run in headless mode",
                    value=self.config.get("headless", True),
                    help="Browser runs in background (no visible window)"
                )
                
                wait_time = st.slider(
                    "Wait time (seconds)",
                    min_value=3,
                    max_value=10,
                    value=self.config.get("wait_time", 5),
                    help="Time to wait after clicking view button"
                )
                
                max_retries = st.slider(
                    "Max retries",
                    min_value=1,
                    max_value=5,
                    value=self.config.get("max_retries", 3),
                    help="Number of retry attempts on failure"
                )
                
                log_level = st.selectbox(
                    "Log Level",
                    ["DEBUG", "INFO", "WARNING", "ERROR"],
                    index=["DEBUG", "INFO", "WARNING", "ERROR"].index(
                        self.config.get("log_level", "INFO")
                    )
                )
            
            # Form buttons
            col1, col2, col3 = st.columns(3)
            with col1:
                save_config = st.form_submit_button("💾 Save Configuration", use_container_width=True)
            with col2:
                test_config = st.form_submit_button("🔗 Test Connection", use_container_width=True)
            with col3:
                validate_config = st.form_submit_button("✅ Validate", use_container_width=True)
            
            if save_config:
                self.config.update({
                    "crm_url": crm_url,
                    "username": username,
                    "password": password,
                    "headless": headless,
                    "wait_time": wait_time,
                    "max_retries": max_retries,
                    "log_level": log_level
                })
                
                if self.save_config():
                    st.success("Configuration saved successfully!")
                    self.add_log("Configuration saved")
            
            if test_config:
                with st.spinner("Testing connection..."):
                    time.sleep(2)  # Simulate connection test
                    st.info("Connection test would be implemented with actual CRM credentials")
            
            if validate_config:
                self.validate_and_show_config()

    def render_file_management(self):
        """Render file management tab"""
        st.subheader("📁 File Management")
        
        # Excel File Upload
        st.markdown("### Upload Excel File")
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Excel file containing Service Order numbers"
        )
        
        if uploaded_file is not None:
            try:
                # Save uploaded file
                file_path = f"uploads/{uploaded_file.name}"
                Path("uploads").mkdir(exist_ok=True)
                
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                self.config["excel_path"] = file_path
                self.save_config()
                
                # Preview data
                df = pd.read_excel(file_path)
                
                st.success(f"✅ File uploaded successfully: {uploaded_file.name}")
                
                # Show preview
                with st.expander("📊 Preview Data"):
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    st.write(f"**File Info:**")
                    st.write(f"• Rows: {len(df)}")
                    st.write(f"• Columns: {len(df.columns)}")
                    st.write(f"• Columns: {', '.join(df.columns.tolist())}")
                
                # Validate Service Order column
                so_columns = [col for col in df.columns if any(
                    keyword in str(col).lower() 
                    for keyword in ['service order', 'so', 'order no', 'job no']
                )]
                
                if so_columns:
                    st.success(f"✅ Found Service Order column: **{so_columns[0]}**")
                    so_count = df[so_columns[0]].dropna().count()
                    st.info(f"Found {so_count} Service Order numbers")
                    
                    # Show sample SO numbers
                    with st.expander("Sample Service Orders"):
                        sample_sos = df[so_columns[0]].dropna().head(10).tolist()
                        for so in sample_sos:
                            st.write(f"• {so}")
                else:
                    st.warning("⚠️ No Service Order column found. Looking for columns with: 'Service Order', 'SO', 'Order No', 'Job No'")
                    
                # Create template option
                if st.button("📝 Create Template"):
                    self.create_template()
                
            except Exception as e:
                st.error(f"Error reading file: {e}")
        
        st.divider()
        
        # Download Management
        st.markdown("### 📥 Download Management")
        
        if os.path.exists(self.download_dir) and os.listdir(self.download_dir):
            files = os.listdir(self.download_dir)
            invoice_files = [f for f in files if f.endswith(('.png', '.jpg', '.jpeg', '.pdf'))]
            
            if invoice_files:
                st.info(f"Found {len(invoice_files)} invoice files")
                
                # Show file list
                with st.expander("📄 View Downloaded Files"):
                    for file in invoice_files[:20]:  # Show first 20
                        st.write(f"• {file}")
                
                # Download all option
                if st.button("📦 Download All as ZIP"):
                    self.download_all_invoices()
            else:
                st.info("No invoice files found in download directory")
        else:
            st.info("Download directory is empty")

    def render_automation(self):
        """Render automation tab"""
        st.subheader("🚀 Run Automation")
        
        # Validation check
        config_status = self.validate_config_silent()
        
        if not config_status['valid']:
            st.error("❌ Configuration incomplete")
            with st.expander("Missing items"):
                for error in config_status['errors']:
                    st.write(f"• {error}")
            st.info("Please complete the configuration in the Configuration tab first.")
            return
        
        # Show current config summary
        with st.expander("📋 Current Configuration", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                st.write("**CRM Details**")
                st.write(f"• URL: {self.config.get('crm_url', 'Not set')[:50]}...")
                st.write(f"• Username: {self.config.get('username', 'Not set')}")
                st.write(f"• Headless: {'Yes' if self.config.get('headless') else 'No'}")
            
            with col2:
                st.write("**File Details**")
                excel_path = self.config.get('excel_path', 'Not set')
                st.write(f"• Excel File: {os.path.basename(excel_path) if excel_path != 'Not set' else 'Not set'}")
                if os.path.exists(excel_path):
                    try:
                        df = pd.read_excel(excel_path)
                        so_columns = [col for col in df.columns if 'service' in str(col).lower() or 'order' in str(col).lower()]
                        if so_columns:
                            so_count = df[so_columns[0]].dropna().count()
                            st.write(f"• Service Orders: {so_count}")
                    except:
                        pass
        
        # Run Automation Section
        st.markdown("### ⚙️ Automation Controls")
        
        if not st.session_state.automation_running:
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if st.button("▶️ Start Automation", use_container_width=True, type="primary"):
                    self.start_automation()
            
            with col2:
                if st.button("🔍 Test Run (1st 5)", use_container_width=True):
                    self.test_automation()
            
            with col3:
                if st.button("🔄 Reset Stats", use_container_width=True):
                    st.session_state.stats = {
                        'total': 0,
                        'success': 0,
                        'failed': 0,
                        'start_time': None,
                        'end_time': None
                    }
                    st.rerun()
            
            # Progress placeholder (will be updated during automation)
            self.progress_placeholder = st.empty()
            self.status_placeholder = st.empty()
            self.logs_placeholder = st.empty()
        
        else:
            # Automation running view
            st.warning("🔄 Automation is currently running...")
            
            # Cancel button
            if st.button("⏹️ Cancel Automation", type="secondary"):
                st.session_state.automation_running = False
                self.add_log("Automation cancelled by user")
                st.rerun()
            
            # Live progress
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Simulate progress (in real app, this would update from the automation thread)
            if st.session_state.current_step:
                status_text.text(f"Current: {st.session_state.current_step}")
            
            # Live logs
            with st.expander("📋 Live Logs", expanded=True):
                if st.session_state.automation_logs:
                    for log in st.session_state.automation_logs[-10:]:
                        st.text(log)

    def render_logs_results(self):
        """Render logs and results tab"""
        st.subheader("📋 Logs & Results")
        
        # Logs Section
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.markdown("### 🔍 Activity Logs")
        with col2:
            if st.button("🗑️ Clear Logs", use_container_width=True):
                st.session_state.automation_logs = []
                st.rerun()
        
        # Display logs
        if st.session_state.automation_logs:
            logs_text = "\n".join(st.session_state.automation_logs)
            st.text_area("Logs", logs_text, height=300, disabled=True)
        else:
            st.info("No logs available")
        
        st.divider()
        
        # Results Section
        st.markdown("### 📊 Results")
        
        if st.session_state.stats['start_time']:
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total Processed", st.session_state.stats['total'])
            
            with col2:
                success_rate = (st.session_state.stats['success'] / st.session_state.stats['total'] * 100) if st.session_state.stats['total'] > 0 else 0
                st.metric("Success Rate", f"{success_rate:.1f}%")
            
            with col3:
                st.metric("Failed", st.session_state.stats['failed'])
            
            # Duration
            if st.session_state.stats['end_time']:
                duration = st.session_state.stats['end_time'] - st.session_state.stats['start_time']
                st.write(f"**Duration:** {duration}")
            
            # Processed Service Orders
            if st.session_state.processed_so:
                with st.expander("✅ Processed Service Orders"):
                    for so in st.session_state.processed_so[:50]:  # Show first 50
                        st.write(f"• {so}")

    def validate_config_silent(self):
        """Validate configuration without showing messages"""
        errors = []
        
        # Check required fields
        if not self.config.get("crm_url"):
            errors.append("CRM URL is required")
        if not self.config.get("username"):
            errors.append("Username is required")
        if not self.config.get("password"):
            errors.append("Password is required")
        
        # Check Excel file
        excel_path = self.config.get("excel_path")
        if not excel_path:
            errors.append("Excel file is required")
        elif not os.path.exists(excel_path):
            errors.append(f"Excel file not found: {excel_path}")
        
        return {"valid": len(errors) == 0, "errors": errors}

    def validate_and_show_config(self):
        """Validate configuration and show results"""
        errors = []
        warnings = []
        
        # Validate CRM URL
        url = self.config.get("crm_url", "")
        if not url:
            errors.append("CRM URL is required")
        elif not url.startswith(("http://", "https://")):
            warnings.append("CRM URL should start with http:// or https://")
        
        # Validate credentials
        if not self.config.get("username"):
            errors.append("Username is required")
        if not self.config.get("password"):
            errors.append("Password is required")
        
        # Validate Excel file
        excel_path = self.config.get("excel_path")
        if not excel_path:
            errors.append("Excel file is required")
        elif not os.path.exists(excel_path):
            errors.append(f"Excel file not found: {excel_path}")
        else:
            try:
                df = pd.read_excel(excel_path)
                if len(df) == 0:
                    warnings.append("Excel file is empty")
                
                # Check for SO number column
                so_columns = [col for col in df.columns if 'service' in str(col).lower() or 'order' in str(col).lower() or 'so' in str(col).lower()]
                if not so_columns:
                    warnings.append("No column name containing 'Service Order' or 'SO' found")
            except Exception as e:
                errors.append(f"Invalid Excel file: {e}")
        
        # Show results
        if errors:
            st.error("❌ Configuration Errors")
            for error in errors:
                st.write(f"• {error}")
        
        if warnings:
            st.warning("⚠️ Configuration Warnings")
            for warning in warnings:
                st.write(f"• {warning}")
        
        if not errors and not warnings:
            st.success("✅ Configuration is valid!")

    def create_template(self):
        """Create Excel template"""
        try:
            template_data = {
                'Service Order no': ['SO001', 'SO002', 'SO003'],
                'Customer Name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
                'Date': ['2024-01-01', '2024-01-02', '2024-01-03'],
                'Amount': [100.00, 150.50, 200.00]
            }
            
            df = pd.DataFrame(template_data)
            
            # Convert to bytes for download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Service Orders')
            
            output.seek(0)
            
            st.download_button(
                label="📥 Download Template",
                data=output,
                file_name="CRM_Invoice_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Failed to create template: {e}")

    def download_all_invoices(self):
        """Create and download ZIP of all invoices"""
        try:
            if os.path.exists(self.download_dir):
                files = os.listdir(self.download_dir)
                invoice_files = [f for f in files if f.endswith(('.png', '.jpg', '.jpeg', '.pdf'))]
                
                if invoice_files:
                    # Create ZIP in memory
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for file in invoice_files:
                            file_path = os.path.join(self.download_dir, file)
                            zip_file.write(file_path, file)
                    
                    zip_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download Invoices ZIP",
                        data=zip_buffer,
                        file_name="invoices.zip",
                        mime="application/zip"
                    )
                else:
                    st.warning("No invoice files found to download")
            else:
                st.warning("Download directory doesn't exist")
                
        except Exception as e:
            st.error(f"Failed to create ZIP: {e}")

    def run_automation_thread(self):
        """Run automation in background thread"""
        try:
            # Import and run main automation
            sys.path.append('.')
            from main import CRMAutomation
            
            # Update session state
            st.session_state.automation_running = True
            st.session_state.stats['start_time'] = datetime.now()
            
            # Load Excel file to get total count
            try:
                df = pd.read_excel(self.config["excel_path"])
                so_columns = [col for col in df.columns if 'service' in str(col).lower() or 'order' in str(col).lower()]
                if so_columns:
                    so_count = df[so_columns[0]].dropna().count()
                    st.session_state.stats['total'] = so_count
            except:
                pass
            
            # Create and run automation
            automation = CRMAutomation(
                config_path=self.config_file,
                excel_path=self.config["excel_path"]
            )
            
            # Run the automation
            automation.run()
            
            # Update stats from automation
            st.session_state.stats['success'] = automation.stats.get('success', 0)
            st.session_state.stats['failed'] = automation.stats.get('failed', 0)
            st.session_state.stats['end_time'] = datetime.now()
            
            # Add completion log
            self.add_log(f"Automation completed! Success: {automation.stats['success']}/{automation.stats['total']}")
            
        except Exception as e:
            self.add_log(f"Automation error: {e}", "ERROR")
        finally:
            st.session_state.automation_running = False

    def start_automation(self):
        """Start the automation process"""
        # Start in background thread
        thread = threading.Thread(target=self.run_automation_thread, daemon=True)
        thread.start()
        
        st.session_state.automation_running = True
        self.add_log("Automation started")
        
        # Show info message
        st.info("Automation started in background. Check the Logs tab for progress.")

    def test_automation(self):
        """Run test automation on first 5 orders"""
        self.add_log("Test automation started (first 5 orders)")
        # Similar to start_automation but with limit
        # Implementation would be similar but with limit parameter

    def run(self):
        """Main entry point for Streamlit app"""
        self.render_sidebar()
        self.render_main_content()

# Run the app
if __name__ == "__main__":
    app = StreamlitCRMAutomation()
    app.run()