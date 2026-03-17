import os
import json
import asyncio
from typing import TypedDict, List

from langgraph.prebuilt import create_react_agent
from langchain_openai import ChatOpenAI
from langchain_core.tools import tool

from mcp import ClientSession
from mcp.client.sse import sse_client

# -----------------------------
# CONFIG
# -----------------------------
GITMCP_URL = "https://gitmcp.io/HariniBrillersys/Data"
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

if not OPENAI_API_KEY:
    print("WARNING: OPENAI_API_KEY environment variable is not set. Please set it in your .env file.")

# -----------------------------
# MCP TOOL WRAPPER
# -----------------------------
# We will create a generic tool that the LLM can use to call MCP tools.
# We establish a new connection each time for simplicity, but in a 
# long-running app, you'd keep the session open.

@tool
def call_mcp_repository_tool(tool_name: str, arguments_json: str) -> str:
    """
    Call an MCP tool on the repository by its exact name and a JSON-formatted arguments string.
    """
    async def run_tool():
        async with sse_client(GITMCP_URL) as streams:
            async with ClientSession(*streams) as session:
                await session.initialize()
                
                try:
                    args = json.loads(arguments_json)
                except json.JSONDecodeError:
                    return "Error: arguments_json must be a valid JSON string."
                    
                result = await session.call_tool(tool_name, args)
                
                if result.isError:
                    return f"Error from tool: {result.content}"
                
                if isinstance(result.content, list):
                    texts = []
                    for c in result.content:
                        try:
                            texts.append(c.text)
                        except AttributeError:
                            texts.append(str(c))
                    # Return truncated output to avoid context length issues with massive file results
                    full_text = "\n".join(texts)
                    return full_text[:4000] + "\n...(truncated if too long)"
                return str(result.content)

    return asyncio.run(run_tool())


# -----------------------------
# FETCH AVAILABLE TOOLS & RUN AGENT
# -----------------------------
async def main():
    print("Connecting to MCP server to fetch available tools...")
    
    # Connect to MCP to discover available tools and schemas
    async with sse_client(GITMCP_URL) as streams:
        async with ClientSession(*streams) as session:
            await session.initialize()
            tools_response = await session.list_tools()
            
            tool_descriptions = []
            for t in tools_response.tools:
                schema_str = json.dumps(t.inputSchema)
                tool_descriptions.append(f"- Tool '{t.name}': {t.description}\n  Schema: {schema_str}")
                
            available_tools_str = "\n".join(tool_descriptions)

    # Initialize LLM & Agent
    llm = ChatOpenAI(
        model="gpt-4o-mini",
        api_key=OPENAI_API_KEY,
        temperature=0
    )

    # The system prompt gives the agent all the info it needs about the tools
    system_prompt = f"""You are a helpful software engineering assistant using an MCP repository server.
The server has the following specific tools:

{available_tools_str}

If the user asks you to do something (e.g., 'list files'), and there isn't an exact tool for it, try using the available search tools (like `search_Data_code` with a broad query like '*' or 'README') to find the structure. Or, inform the user about what tools ARE available instead. 

To call a tool, use `call_mcp_repository_tool`. Provide the precise `tool_name` and provide `arguments_json` strictly matching the schema for that tool.
"""

    agent_executor = create_react_agent(
        llm,
        tools=[call_mcp_repository_tool],
        state_modifier=system_prompt  # Pass the prompt using the proper argument for LangGraph's prebuilt React agent
    )

    print("\n--- Agent Initialized ---\n")
    
    while True:
        try:
            query = input("Ask about the repo (or 'quit'): ")
            if query.lower() in ("quit", "exit"):
                break
                
            print("Agent is thinking (this may take a few seconds)...")
            
            # Run the agent
            for event in agent_executor.stream({"messages": [("user", query)]}):
                # Print tool calls and results as they happen
                for val in event.values():
                    messages = val.get("messages", [])
                    if messages:
                        message = messages[-1]
                        if message.type == "ai" and message.tool_calls:
                            print(f"\n[Agent Calling Tool]: {message.tool_calls[0]['args']}\n")
                        elif message.type == "tool":
                            print(f"[Tool Returned {len(message.content)} characters]")

            # The final answer is the last AI message
            final_state = agent_executor.invoke({"messages": [("user", query)]})
            final_message = final_state["messages"][-1].content
            
            print("\nAgent Answer:\n" + "-"*40)
            print(final_message)
            print("-" * 40 + "\n")
            
        except KeyboardInterrupt:
            break

if __name__ == "__main__":
    asyncio.run(main())