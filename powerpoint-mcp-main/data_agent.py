import os
from typing import TypedDict
from dotenv import load_dotenv

from langchain_core.messages import HumanMessage
from langchain_openai import ChatOpenAI
from langgraph.graph import StateGraph, END

# Load environment variables (e.g., OPENAI_API_KEY)
load_dotenv()

# -----------------------------
# Agent State
# -----------------------------
class AgentState(TypedDict):
    query: str
    csv_data: str
    answer: str

# -----------------------------
# Tool / Node: Analyze Data
# -----------------------------
def analyze_data(state: AgentState):
    """
    Uses an LLM to analyze the provided CSV data based on the user's query.
    """
    print("[analyze_data] Analyzing the CSV data with LLM...")
    
    query = state.get("query", "")
    csv_data = state.get("csv_data", "")
    
    # Initialize LLM
    llm = ChatOpenAI(
        model="gpt-4o-mini",
        openai_api_key=os.getenv("OPENAI_API_KEY")
    ).bind(response_format={"type": "json_object"})

    prompt = f"""
You are an expert Data Analysis Agent. Your task is to analyze the provided dataset and answer the user's question.
Crucially, you must output your response as a valid JSON object suitable for generating a PowerPoint presentation.

Your core operations are:
1. Data Parsing: Understand the structure, columns, and data types of the provided dataset.
2. Data Profiling: Identify missing values, anomalies, unique categorical values, and general distributions.
3. Statistical Analysis: Calculate relevant metrics (mean, median, standard deviation, correlations, etc.) based on the user's question.

The JSON output MUST follow this exact structure:
{{
  "presentation": {{
    "title": "<A concise title for the entire report based on the user's query>",
    "slides": [
      {{
        "title": "<Slide title>",
        "content": [
          "<Bullet point 1 conveying an insight or metric>", 
          "<Bullet point 2 conveying an insight or metric>"
        ],
        "speaker_notes": "<Detailed explanation or insights derived from the data for this slide>"
      }}
    ]
  }}
}}

Ensure that the output is ONLY valid JSON. If the data doesn't contain the answer, return a JSON with a single slide stating that there is not enough information based on the provided data.

CSV Data:
{csv_data}

Question:
{query}
"""

    response = llm.invoke([HumanMessage(content=prompt)])
    
    # Update state with the final answer
    return {"answer": response.content}

# -----------------------------
# Build the LangGraph Workflow
# -----------------------------
def create_data_agent_workflow():
    builder = StateGraph(AgentState)

    # Add node and execution flow
    builder.add_node("analyze_data", analyze_data)
    builder.set_entry_point("analyze_data")
    builder.add_edge("analyze_data", END)

    # Compile the graph
    return builder.compile()

# Automatically compile for easy import
graph = create_data_agent_workflow()
