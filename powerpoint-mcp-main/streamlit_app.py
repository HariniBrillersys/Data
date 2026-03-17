import streamlit as st
import pandas as pd
from data_agent import graph

st.set_page_config(page_title="LangGraph Data Agent", layout="wide")

st.title("📊 LangGraph Data Agent")
st.write("Upload a CSV file and ask questions about your data.")

# Upload CSV
uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

if uploaded_file is not None:
    # Read the CSV to show a preview to the user
    df = pd.read_csv(uploaded_file)
    
    st.subheader("Data Preview")
    st.dataframe(df.head(10))  # Preview the first 10 rows
    
    # Convert a subset of the dataframe to CSV string format to pass to our LLM
    # We restrict it to the first 100 rows so it doesn't exceed LLM token limits
    # In a more advanced implementation, you might want to use a PandasAgent or a code execution node
    csv_string = df.head(100).to_csv(index=False)
    
    # User Query Input
    st.subheader("Ask the Agent")
    query = st.text_input("Enter your question about this data:")
    
    if st.button("Analyze Data"):
        if query:
            with st.spinner("Analyzing data through LangGraph..."):
                # Pass the query and the stringified CSV data into the langgraph State
                initial_state = {
                    "query": query,
                    "csv_data": csv_string
                }
                
                # Execute the LangGraph workflow
                result = graph.invoke(initial_state)
                answer_json_str = result.get("answer", "{}")
                
                import json
                try:
                    parsed_answer = json.loads(answer_json_str)
                except json.JSONDecodeError:
                    parsed_answer = answer_json_str

                # Store result in session state to survive app reruns
                st.session_state['parsed_answer'] = parsed_answer
        else:
            st.warning("Please enter a question to analyze.")

    # Check if we have an answer in the session state to display
    if 'parsed_answer' in st.session_state:
        st.success("Analysis Complete!")
        parsed_answer = st.session_state['parsed_answer']

        if isinstance(parsed_answer, dict):
            st.write("### Extracted Insights")
            if "presentation" in parsed_answer and "title" in parsed_answer["presentation"]:
                st.markdown(f"**{parsed_answer['presentation']['title']}**")
            st.json(parsed_answer)
            
            # Generate PPT Button block
            st.markdown("---")
            st.subheader("Generate Presentation")
            st.write("Click below to create a PowerPoint deck from these insights.")
            
            if st.button("Generate PPTX", key="gen_ppt"):
                with st.spinner("Generating PowerPoint presentation..."):
                    from ppt_generator import generate_ppt_from_json
                    import os
                    
                    try:
                        os.makedirs("outputs", exist_ok=True)
                        output_file = "outputs/generated_report.pptx"
                        ppt_path = generate_ppt_from_json(parsed_answer, output_filename="generated_report.pptx")
                        
                        st.session_state['generated_ppt_path'] = output_file
                        st.success(f"Presentation generated successfully!")
                    except Exception as e:
                        st.error(f"Failed to generate presentation: {e}")
            
            # Show download button if PPT has been generated
            if 'generated_ppt_path' in st.session_state:
                import os
                output_file = st.session_state['generated_ppt_path']
                if os.path.exists(output_file):
                    with open(output_file, "rb") as file:
                        st.download_button(
                            label="📥 Download Presentation",
                            data=file,
                            file_name="Data_Analysis_Report.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
        else:
            st.write("### Answer (Raw)")
            st.write(parsed_answer)
