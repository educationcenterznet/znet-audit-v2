try:
    # 嘗試匯入 pptx。如果失敗，則執行安裝。
    import pptx
except ImportError:
    st.warning("Force installing 'python-pptx' package due to environment error. Please wait...")
    try:
        # 強制使用 pip 執行安裝指令
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-pptx"])
        import pptx # 再次嘗試匯入
        st.success("Installation successful! Restarting the application...")
        # 由於套件被動態安裝，最好的做法是強制 Streamlit 重新載入
        st.experimental_rerun() 
    except Exception as e:
        st.error(f"FATAL ERROR during force installation: {e}")
        st.stop()
# ------------------------------------------------------------------

# 現在 pptx 已經確定安裝或存在，可以安全地匯入後續的模組
import backend_logic as bl  # 這裡不再會報錯
import json
import zipfile
import io

# Setting page title and layout
st.set_page_config(page_title="Zyxel Training AI Auditor", layout="wide")
st.title("Zyxel Training Content Automated Audit System")

# Sidebar: API Key Configuration
with st.sidebar:
    st.header("Configuration")
    api_key = st.text_input("Enter Gemini API Key", type="password")
    # Recommended model switch to avoid 429 quota errors
    model_choice = st.radio(
        "Select AI Model:",
        ('gemini-2.5-flash', 'gemini-2.5-pro'),
        index=0, # Default to flash
        help="Flash offers significantly higher quota/speed, Pro offers maximum accuracy (lower quota)."
    )

    if not api_key:
        st.warning("Please enter your API Key to enable AI features.")

# Use Tabs to separate the three main workflows
tab1, tab2, tab3 = st.tabs(["1. Zycamp Source Analysis", "2. ZCNE Course Indexing", "3. Audit Report Generation"])

# ==========================================
# TAB 1: Zycamp Source Processing
# ==========================================
with tab1:
    st.header("Step 1: Process New Feature Presentation (Source)")
    
    # 1. Upload PPT
    uploaded_zycamp = st.file_uploader("Upload Zycamp PPT (Single File)", type=['pptx'], key="zycamp_ppt")
    
    if uploaded_zycamp:
        # Get the base filename for dynamic naming
        ppt_base_name = os.path.splitext(uploaded_zycamp.name)[0]
        
        # Button: Extract Text
        if st.button("Extract Text", key="btn_extract_zycamp"):
            with st.spinner("Extracting text and creating structured content..."):
                text_content = bl.extract_text_from_ppt_stream(uploaded_zycamp)
                st.session_state['zycamp_txt'] = text_content
                st.success("Text extraction complete!")
        
        # Display Download Button for TXT
        if 'zycamp_txt' in st.session_state:
            st.download_button(
                label="Download TXT Source Content",
                data=st.session_state['zycamp_txt'],
                file_name=f"{ppt_base_name}_Source_Content.txt",
                mime="text/plain"
            )
            
            st.divider()
            
            # Button: AI Analysis (Generate JSON)
            if st.button("AI Analyze & Generate JSON (Source List)", key="btn_analyze_zycamp"):
                if not api_key:
                    st.error("Please enter your API Key in the sidebar.")
                else:
                    with st.spinner(f"AI is analyzing new features using {model_choice}..."):
                        json_result = bl.call_gemini_api(
                            bl.PROMPT_SOURCE_ANALYSIS, 
                            st.session_state['zycamp_txt'], 
                            api_key,
                            model_name=model_choice,
                            output_json=True
                        )
                        if "API Error" in json_result:
                            st.error(f"AI Analysis Failed: {json_result}")
                        else:
                            st.session_state['zycamp_json'] = json_result
                            st.success("Analysis complete! Source List JSON generated.")
            
            # Display JSON Download Button
            if 'zycamp_json' in st.session_state:
                st.download_button(
                    label="Download Source_List.json",
                    data=st.session_state['zycamp_json'],
                    file_name=f"{ppt_base_name}_Source_List.json",
                    mime="application/json"
                )

# ==========================================
# TAB 2: ZCNE Course Indexing
# ==========================================
with tab2:
    st.header("Step 2: Build Core Curriculum Index (Targets)")
    
    # 1. Upload multiple PPTs
    uploaded_zcne_ppts = st.file_uploader("Upload ZCNE Course PPTs (Multi-select)", type=['pptx'], accept_multiple_files=True, key="zcne_ppts")
    
    if uploaded_zcne_ppts:
        if st.button("Batch Process & Generate Indices", key="btn_process_zcne"):
            if not api_key:
                st.error("Please enter your API Key in the sidebar.")
            else:
                
                # --- START: Processing and Zipping ---
                
                # Clear previous session data for this tab
                st.session_state['zcne_txt_zip_data'] = None
                st.session_state['zcne_json_zip_data'] = None
                st.session_state['zcne_files_count'] = 0
                
                # Prepare Zip file buffers
                zip_buffer_txt = io.BytesIO()
                zip_buffer_json = io.BytesIO()
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                txt_files_map = {}
                json_files_map = {}

                with st.spinner(f"Processing {len(uploaded_zcne_ppts)} files and building indices using {model_choice}..."):
                    for i, ppt_file in enumerate(uploaded_zcne_ppts):
                        file_name_base = os.path.splitext(ppt_file.name)[0]
                        status_text.text(f"Processing: {file_name_base}...")
                        
                        # A. Extract Text
                        txt_content = bl.extract_text_from_ppt_stream(ppt_file)
                        txt_files_map[f"{file_name_base}_Content.txt"] = txt_content
                        
                        # B. AI Indexing
                        json_result = bl.call_gemini_api(
                            bl.PROMPT_TARGET_INDEXING,
                            txt_content,
                            api_key,
                            model_name=model_choice,
                            output_json=True
                        )

                        if "API Error" in json_result:
                            st.error(f"AI Indexing Failed for {file_name_base}: {json_result}")
                        else:
                            json_files_map[f"Target_Index_{file_name_base}.json"] = json_result
                        
                        # Update progress bar
                        progress_bar.progress((i + 1) / len(uploaded_zcne_ppts))

                # Package TXT Zip
                with zipfile.ZipFile(zip_buffer_txt, "w") as zf:
                    for fname, content in txt_files_map.items():
                        zf.writestr(fname, content.encode('utf-8'))
                
                # Package JSON Zip
                with zipfile.ZipFile(zip_buffer_json, "w") as zf:
                    for fname, content in json_files_map.items():
                        zf.writestr(fname, content.encode('utf-8'))

                # --- 關鍵修正: 儲存 Zip 內容到 Session State ---
                st.session_state['zcne_txt_zip_data'] = zip_buffer_txt.getvalue()
                st.session_state['zcne_json_zip_data'] = zip_buffer_json.getvalue()
                st.session_state['zcne_files_count'] = len(json_files_map)
                
                st.success("All files processed successfully! Download buttons are now available.")
        
        # --- 關鍵修正: 將下載按鈕移到 if st.button 外面，並檢查 Session State ---
        if st.session_state.get('zcne_json_zip_data') and st.session_state.get('zcne_txt_zip_data'):
            files_count = st.session_state['zcne_files_count']
            
            st.divider() # Separation line after processing success

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label=f"Download {files_count} Module TXT Files (Zip)",
                    data=st.session_state['zcne_txt_zip_data'],
                    file_name="ZCNE_Modules_TXT.zip",
                    mime="application/zip"
                )
            with col2:
                st.download_button(
                    label=f"Download {files_count} Target Index JSONs (Zip)",
                    data=st.session_state['zcne_json_zip_data'],
                    file_name="ZCNE_Target_Indices.zip",
                    mime="application/zip"
                )

# ==========================================
# TAB 3: Report Generation
# ==========================================
with tab3:
    st.header("Step 3: Generate Final Audit Reports (Mapping)")
    
    col1, col2 = st.columns(2)
    with col1:
        source_file = st.file_uploader("Upload Source_List.json (from Step 1)", type=['json'])
    with col2:
        target_files = st.file_uploader("Upload Target_Index.json (Multi-select, from Step 2)", type=['json'], accept_multiple_files=True)
        
    if source_file and target_files:
        if st.button("Generate All Audit Reports", key="btn_report"):
            if not api_key:
                st.error("Please enter your API Key in the sidebar.")
            else:
                # Read Source List
                source_content = source_file.getvalue().decode("utf-8")
                
                # Prepare Report Zip
                zip_buffer_report = io.BytesIO()
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                reports_generated = 0
                
                with st.spinner(f"AI is cross-referencing and generating reports using {model_choice}..."):
                    for i, target_file in enumerate(target_files):
                        # Use the JSON filename base for the MD report
                        json_file_base = os.path.splitext(target_file.name)[0] 
                        report_filename = json_file_base.replace("Target_Index_", "Report_") + ".md"
                        
                        status_text.text(f"Generating report: {report_filename}...")
                        
                        target_content = target_file.getvalue().decode("utf-8")
                        
                        # Combine Prompt 3 (inserting the two JSONs)
                        mapping_prompt = bl.PROMPT_MAPPING_REPORT.format(
                            source_json=source_content,
                            target_json=target_content
                        )
                        
                        # Call AI (non-JSON output)
                        report_content = bl.call_gemini_api(
                            "", # Prompt content is already in mapping_prompt
                            mapping_prompt,
                            api_key,
                            model_name=model_choice,
                            output_json=False
                        )
                        
                        if "API Error" in report_content:
                            st.error(f"Report Generation Failed for {report_filename}: {report_content}")
                        else:
                            # Add to Zip
                            with zipfile.ZipFile(zip_buffer_report, "a") as zf:
                                zf.writestr(report_filename, report_content.encode('utf-8'))
                            reports_generated += 1
                            
                        progress_bar.progress((i + 1) / len(target_files))
                
                # --- 關鍵修正: 儲存報告 Zip 內容到 Session State ---
                st.session_state['audit_report_zip_data'] = zip_buffer_report.getvalue()
                st.session_state['reports_generated_count'] = reports_generated
                
                st.success(f"{reports_generated} Reports generated successfully!")
        
        # --- 關鍵修正: 將下載按鈕移到 if st.button 外面，並檢查 Session State ---
        if st.session_state.get('audit_report_zip_data'):
            st.divider()
            reports_count = st.session_state['reports_generated_count']
            st.download_button(
                label=f"Download All {reports_count} Audit Reports (Zip)",
                data=st.session_state['audit_report_zip_data'],
                file_name="Audit_Reports.zip",
                mime="application/zip"

            )
