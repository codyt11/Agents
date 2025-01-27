from huggingface import generate_summary
from composio_crewai import ComposioToolSet, Action
from crewai import Crew, Agent, Task, Process
from langchain_openai import ChatOpenAI
from pathlib import Path
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()

# Summarize the document using the first script
input_doc_path = "./Amarin_CEA_Design_Specification.docx"
summary_output_path = "./summary.docx"
summarized_text = generate_summary(input_doc_path, summary_output_path)

# Ensure the summarization was successful
if summarized_text:
    print(f"Summarized Text: {summarized_text[:300]}...")  # Print a preview of the summary

    # Initialize the language model
    llm = ChatOpenAI(model='gpt-4o')

    # Set up Composio ToolSet
    composio_toolset = ComposioToolSet(output_dir=Path("./ppts"))
    tools = composio_toolset.get_tools(actions=[
        Action.CODEINTERPRETER_EXECUTE_CODE,
        Action.CODEINTERPRETER_GET_FILE_CMD,
        Action.CODEINTERPRETER_RUN_TERMINAL_CMD,
    ])

    # Agent 1: Teaching Assistant
    teaching_assistant = Agent(
        role="Teaching Assistant",
        goal="Help the user/student understand the document",
        backstory=f"""
            You are a teaching assistant with complete knowledge of the document:
            {summarized_text}
            Your job is to answer questions clearly and thoroughly.
        """,
        tools=tools,
    )

    # Task 1: Answer Questions
    answer_questions = Task(
        description=f"""
            Use the summarized document to answer any user questions. Compare the summary with the original document
            to ensure completeness and clarity.
        """,
        expected_output="An accurate and complete answer to the user's question.",
        tools=tools,
        agent=teaching_assistant,
        verbose=True,
    )

    # Execute the task
    crew = Crew(agents=[teaching_assistant], tasks=[answer_questions], process=Process.sequential)
    question_response = crew.kickoff()
    print(f"Question answered: {question_response}")
else:
    print("Failed to summarize the document.")
