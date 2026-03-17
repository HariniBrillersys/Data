import requests
from langchain_core.messages import HumanMessage
from langchain_openai import ChatOpenAI
from langgraph.graph import StateGraph, END
from typing import TypedDict
import os
from dotenv import load_dotenv

load_dotenv()


# -----------------------------
# Config
# -----------------------------
GITMCP_URL = "https://gitmcp.io/HariniBrillersys/Data"


# -----------------------------
# Agent State
# -----------------------------
class AgentState(TypedDict):
    query: str
    repo_data: str
    answer: str


# -----------------------------
# Tool: Fetch data from GitMCP
# -----------------------------
def fetch_repo_data(state: AgentState):

    query = state["query"]

    response = requests.get(GITMCP_URL)

    if response.status_code == 200:
        data = response.text
    else:
        data = "Unable to fetch repo data."

    return {"repo_data": data}


# -----------------------------
# Data Agent
# -----------------------------


import os

llm = ChatOpenAI(
    model="gpt-4o-mini",
    openai_api_key=os.getenv("OPENAI_API_KEY")
)


def data_agent(state: AgentState):

    query = state["query"]
    repo_data = state["repo_data"]

    prompt = f"""
You are a data agent.

Use the following repository data to answer the question.

Repository Data:
{repo_data}

Question:
{query}
"""

    response = llm.invoke([HumanMessage(content=prompt)])

    return {"answer": response.content}


# -----------------------------
# LangGraph Workflow
# -----------------------------
builder = StateGraph(AgentState)

builder.add_node("fetch_repo", fetch_repo_data)
builder.add_node("data_agent", data_agent)

builder.set_entry_point("fetch_repo")

builder.add_edge("fetch_repo", "data_agent")
builder.add_edge("data_agent", END)

graph = builder.compile()


# -----------------------------
# Run Agent
# -----------------------------
result = graph.invoke(
    {
        "query": "What files exist in the repo?"
    }
)

print(result["answer"])