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
            # st.success("✅ Word template uploaded successfully!")

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


        # Add search functionality below the tabs
    st.markdown("---")  # Add a separator line
    search_term = st.text_input("🔍 Search for a specific placeholder", "")
    
    if search_term:
        st.write("**Search Results:**")
        
        # Search in Variables
        variables_found = False
        if hasattr(st.session_state, 'generator') and st.session_state.generator.variable_dict:
            filtered_vars = {k: v for k, v in st.session_state.generator.variable_dict.items() 
                        if search_term.lower() in k.lower()}
            if filtered_vars:
                variables_found = True
                # st.write("**From Variables:**")
                for key, value in filtered_vars.items():
                    st.write(f"• **{key}** → {value}")
        
        # Search in Budget
        budget_found = False
        if hasattr(st.session_state, 'generator') and st.session_state.generator.budget_placeholders:
            filtered_budget = {k: v for k, v in st.session_state.generator.budget_placeholders.items() 
                            if search_term.lower() in k.lower()}
            if filtered_budget:
                budget_found = True
                # st.write("**From Budget:**")
                for key, value in filtered_budget.items():
                    st.write(f"• **{key}** → {value}")
        
        # Show warning if nothing found
        if not variables_found and not budget_found:
            st.warning("No matching placeholders found")
    
    # Generate SOW section
    if st.session_state.excel_uploaded and st.session_state.word_uploaded:
        st.markdown("---")
        st.subheader(" Generate Statement of Work")
        
        client_name = st.session_state.replacements.get('{CLIENT NAME}', 'Unknown_Client')
        output_filename = f"SOW_{client_name}_{datetime.now().strftime('%Y%m%d')}.docx"

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
                st.error(f" Error generating SOW: {str(e)}")            
            
    else:
        st.markdown("---")
        st.info("👆 Please upload both Excel file and Word template to proceed")
    
    
    # Instructions
    with st.expander("ℹ️ How to Use This Tool", expanded=False):
        st.markdown("""
        ### Step-by-Step Guide:
        
        1. **Upload Excel File**: Upload your pricing sheet containing:
        - A `Variables` sheet with placeholder mappings
        - A `Budget` sheet with phase costs
        
        2. **Upload Word Template**: Upload your SOW template document with placeholders (e.g., `{CLIENT NAME}`)
        
        3. **Review Mappings**: Check the placeholder mappings to ensure all values are correct
        
        4. **Generate SOW**: Click the "Generate SOW" button to create your document
        
        5. **Download**: Download the generated SOW document
        
        ### Excel File Requirements:
        - **Variables Sheet**: Must have columns "Variable Component" and "Example Value"
        - **Budget Sheet**: Phase names in column A (0), amounts in column F (5)
        
        ### Date Formatting:
        Dates are automatically formatted as `DD-MMM-YY` (e.g., 13-Aug-25)
        """)

    
    
    # Footer
    st.markdown("---")
    st.markdown("*Built with Streamlit | SOW Generator v1.0*")

if __name__ == "__main__":
    main()


