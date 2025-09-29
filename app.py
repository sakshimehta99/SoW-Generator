import streamlit as st
import pandas as pd
from datetime import datetime
import io
import tempfile
import os
from sow_generator import SOWGenerator

def main():
    st.set_page_config(
        page_title="SOW Generator",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("📄 Statement of Work Generator")
    st.markdown("Upload your Excel data and Word template to generate a customized Statement of Work")
    
    # Initialize session state
    if 'replacements' not in st.session_state:
        st.session_state.replacements = {}
    if 'excel_uploaded' not in st.session_state:
        st.session_state.excel_uploaded = False
    if 'word_uploaded' not in st.session_state:
        st.session_state.word_uploaded = False
    
    # Create two columns for file uploads
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Upload Excel File")
        excel_file = st.file_uploader(
            "Select Excel file containing Variables and Budget sheets",
            type=['xlsx', 'xls'],
            help="This file should contain 'Variables' and 'Budget' sheets with placeholder mappings"
        )
        
        if excel_file is not None:
            try:
                # Process Excel file
                with st.spinner("Processing Excel file..."):
                    generator = SOWGenerator()
                    replacements = generator.process_excel_file(excel_file)
                    st.session_state.replacements = replacements
                    st.session_state.generator = generator
                    st.session_state.excel_uploaded = True
                
                st.success("✅ Excel file processed successfully!")
                
                # Display mapping preview
                st.subheader("📋 Placeholder Mapping Preview")
                
                # Create tabs for Variables and Budget
                var_tab, budget_tab = st.tabs(["Variables", "Budget"])
                
                with var_tab:
                    if hasattr(st.session_state, 'generator') and st.session_state.generator.variable_dict:
                        df_vars = pd.DataFrame(list(st.session_state.generator.variable_dict.items()))
                        df_vars.columns = ['Placeholder', 'Value']
                        st.dataframe(df_vars, use_container_width=True)
                    else:
                        st.info("No variable placeholders found")
                
                with budget_tab:
                    if hasattr(st.session_state, 'generator') and st.session_state.generator.budget_placeholders:
                        df_budget = pd.DataFrame(list(st.session_state.generator.budget_placeholders.items()))
                        df_budget.columns = ['Placeholder', 'Value']
                        st.dataframe(df_budget, use_container_width=True)
                    else:
                        st.info("No budget placeholders found")
                
            except Exception as e:
                st.error(f"❌ Error processing Excel file: {str(e)}")
                st.session_state.excel_uploaded = False
    
    with col2:
        st.subheader("📝 Upload Word Template")
        word_file = st.file_uploader(
            "Select Word template document",
            type=['docx'],
            help="This should be your SOW template with placeholders to be replaced"
        )
        
        if word_file is not None:
            st.session_state.word_uploaded = True
            st.success("✅ Word template uploaded successfully!")
    
    # Generate SOW section
    if st.session_state.excel_uploaded and st.session_state.word_uploaded:
        st.markdown("---")
        st.subheader("🚀 Generate Statement of Work")
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            if st.button("📖 Preview Document", type="secondary", use_container_width=True):
                try:
                    with st.spinner("Generating preview..."):
                        generator = SOWGenerator()
                        
                        # Save uploaded files to temporary location
                        if excel_file is not None and word_file is not None:
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
                                tmp_excel.write(excel_file.getvalue())
                                excel_path = tmp_excel.name
                            
                            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_word:
                                tmp_word.write(word_file.getvalue())
                                word_path = tmp_word.name
                            
                            # Generate preview
                            preview_content = generator.generate_preview(excel_path, word_path)
                            
                            # Clean up temp files
                            os.unlink(excel_path)
                            os.unlink(word_path)
                        else:
                            st.error("Files not available. Please upload both files first.")
                            return
                        
                        # Store preview in session state to display below
                        st.session_state.preview_content = preview_content
                
                except Exception as e:
                    st.error(f"❌ Error generating preview: {str(e)}")
        
        with col2:
            client_name = st.session_state.replacements.get('{CLIENT NAME}', 'Unknown_Client')
            output_filename = f"SOW_{client_name}_{datetime.now().strftime('%Y%m%d')}.docx"
            
            # if st.button("⬇️ Generate & Download", type="primary", use_container_width=True):
            #     try:
            #         with st.spinner("Generating SOW document..."):
            #             generator = SOWGenerator()
                        
            # Save uploaded files to temporary location
            if excel_file is not None and word_file is not None:
                try:
                    generator = SOWGenerator()
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
                        tmp_excel.write(excel_file.getvalue())
                        excel_path = tmp_excel.name
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_word:
                        tmp_word.write(word_file.getvalue())
                        word_path = tmp_word.name
                
                    # Generate final document
                    output_buffer = generator.generate_sow(excel_path, word_path)
                    
                    # Clean up temp files
                    os.unlink(excel_path)
                    os.unlink(word_path)
            
            
                # Provide download
                    st.download_button(
                        label="📥 Generate and Download SOW",
                        data=output_buffer.getvalue(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        type="primary"
                    )

                    # st.success(f"✅ SOW generated successfully! Click the download button above to save '{output_filename}'")

                except Exception as e:
                    st.error(f"❌ Error generating SOW: {str(e)}")            
                
            
    

        
        # Display preview full width if available
        if hasattr(st.session_state, 'preview_content') and st.session_state.preview_content:
            st.markdown("---")
            st.subheader("📖 Document Preview")
            st.text_area(
                "Preview Content",
                value=st.session_state.preview_content,
                height=400,
                help="This is a text representation of your generated document",
                key="preview_display"
            )
    
    else:
        st.markdown("---")
        st.info("👆 Please upload both Excel file and Word template to proceed")
    
    # Tips section
    if st.session_state.excel_uploaded and st.session_state.word_uploaded:
        st.markdown("---")
        st.info("💡 **Tips:**\n- Use Preview to verify your document\n- Check placeholder mappings above\n- Ensure all required data is present")
    
    # Footer
    st.markdown("---")
    st.markdown("*Built with Streamlit | SOW Generator v1.0*")

if __name__ == "__main__":
    main()
