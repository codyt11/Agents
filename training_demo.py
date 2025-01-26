from composio_crewai import ComposioToolSet, Action
from crewai import Crew, Agent, Task, Process
from langchain_openai import ChatOpenAI
from dotenv import load_dotenv
from pathlib import Path
from docx import Document
import os
import glob
import re
import shutil

# Load environment variables
load_dotenv()

# Initialize the language model
llm = ChatOpenAI(model='gpt-4o')

# Define input and output file paths
DOC_path = "./Walmart_Stocking_Training_Guide.docx"
DOC_ID = Document(DOC_path)
FINAL_OUTPUT_PPT = './ppts/Walmart_Stocking_Training_Presentation_final.pptx'

# Ensure document exists
if os.path.exists(DOC_path):
    content = [paragraph.text.strip() for paragraph in DOC_ID.paragraphs if paragraph.text.strip()]
    print("Extracted Document Content:", content[:5])  # Print first 5 paragraphs
else:
    raise FileNotFoundError(f"The document at {DOC_path} does not exist or is inaccessible.")

# Set up Composio ToolSet
composio_toolset = ComposioToolSet(output_dir=Path("./ppts"))
tools = composio_toolset.get_tools(actions=[
    Action.CODEINTERPRETER_EXECUTE_CODE,
    Action.CODEINTERPRETER_GET_FILE_CMD,
    Action.CODEINTERPRETER_RUN_TERMINAL_CMD,
])

# Agent 1: PowerPoint Creator
ppt_creator_agent = Agent(
    role="Creative PowerPoint Designer",
    goal="Analyze the Word document and Design a visually stunning and creative PowerPoint presentation",
    backstory=f"""
        You are a creative AI assistant specializing in designing visually appealing PowerPoint presentations using the Python-PPTX library.
        Your task is to analyze the Word document provided: {content}.
        In addition to including the document's key insights, you should:
        - Use custom themes, fonts, and color schemes to make the presentation professional and engaging.
        - Add relevant images, icons, or illustrations where appropriate.
        - Include slide transitions and animations to enhance the overall presentation.
        - Ensure the layout is visually balanced and easy to read.
        Your primary goal is to create a presentation that is not only informative but also captivating and polished.
        save the presentation as a .pptx file
    """,
    tools=tools,
)

# Agent 2: PowerPoint Evaluator
ppt_evaluator_agent = Agent(
    role="PowerPoint Evaluator",
    goal="Evaluate the PowerPoint presentation for relevance and completeness",
    backstory=f"""
        You are an AI evaluator responsible for analyzing a PowerPoint presentation created from a Word document.
        Compare the content of the PowerPoint presentation with the original document: {content}.
        Assess whether all relevant information is included in the presentation.
        If complete, provide approval for saving. Otherwise, provide feedback for improvement.
    """,
    tools=tools,
)

# Feedback loop
MAX_ITERATIONS = 5  # Prevent infinite loop
iteration = 1
feedback = None

while iteration <= MAX_ITERATIONS:
    print(f"\n--- Iteration {iteration} ---\n")
    current_output_ppt = f'./ppts/Walmart_Stocking_Training_Presentation_v{iteration}.pptx'
    print(f"Expected PowerPoint path: {current_output_ppt}")

    # Pass feedback to Agent 1's backstory
    if feedback and "Incomplete feedback" not in feedback:
        ppt_creator_agent.backstory += f"\nEvaluator Feedback: {feedback}\n"
        print(f"Updated Creator Agent Backstory:\n{ppt_creator_agent.backstory}")

    # Task 1: Create PowerPoint
    ppt_task = Task(
        description=f"""
            Create a detailed and professional PowerPoint presentation based on the content of the document: {content}.
            First, retrieve the content, and then pip installs the python-pptx using a code interpreter.
            If there is feedback from the evaluator, incorporate it to improve the presentation.
            Steps:
            1. Analyze the content of the Word document to identify key topics, sections, and insights.
            2. Use the evaluator's feedback: {feedback} (if available) to revise and enhance the slides.
            3. For each major topic or section in the document, create dedicated slides with:
               - Slide Title: Use the section heading or a summary of the topic.
               - Slide Content: Include key bullet points, summaries, or critical data.
               NOTE: if the content is too much for one slide you can either adjust font size or use multiple slides.
            4. Structure the PowerPoint presentation as follows:
               - Title Slide: Include the document's title and purpose.
               - Table of Contents: Summarize the sections/topics covered in the presentation.
               - Content Slides: Add a slide for each major section/key insight in the document.
               - Conclusion Slide: Summarize the key takeaways.
            5. Style the slide, use colors, shapes, shading, or images to make the slides more engaging
            6. Save the PowerPoint presentation as {current_output_ppt}.
        """,
        expected_output=f"A PowerPoint presentation file named {current_output_ppt} should be created and saved as ./ppts/Walmart_Stocking_Training_Presentation_final_v{iteration}.pptx",
        tools=tools,
        agent=ppt_creator_agent,
        verbose=True,
    )
    crew = Crew(agents=[ppt_creator_agent], tasks=[ppt_task], process=Process.sequential)
    creation_response = crew.kickoff()
    print(f"Creation Response: {creation_response}")

    # Check if PowerPoint was created
    saved_files = glob.glob(f"./ppts/CODEINTERPRETER_GET_FILE_CMD_default_*Walmart_Stocking_Training_Presentation_v{iteration}.pptx")

    if saved_files:
        # Get the first matching file
        actual_saved_path = saved_files[0]
        print(f"Found file: {actual_saved_path}")

        # Move or rename it to the expected path
        shutil.move(actual_saved_path, current_output_ppt)
        print(f"File moved to: {current_output_ppt}")
    else:
        raise FileNotFoundError(f"File not found with the expected pattern in ./ppts for iteration {iteration}.")

    # Task 2: Evaluate PowerPoint
    evaluation_task = Task(
        description=f"""
            Evaluate the PowerPoint presentation saved at {current_output_ppt}.
            Steps:
            1. Compare the slides with the original document content: {content}.
            2. Identify any missing or irrelevant information.
            3. Check for appearance, look for overflow text, suggest background color or styling.
            4. If the presentation is complete, approve it for saving.
            5. If incomplete, provide feedback to improve it.
        """,
        expected_output="Evaluation feedback or approval for saving.",
        tools=tools,
        agent=ppt_evaluator_agent,
        verbose=True,
    )
    crew = Crew(agents=[ppt_evaluator_agent], tasks=[evaluation_task], process=Process.sequential)
    evaluation_response = crew.kickoff()
    print(f"Evaluation Response: {evaluation_response}")

    # Extract feedback or approval
    feedback = evaluation_response.result if hasattr(evaluation_response, 'result') else "Incomplete feedback"
    print(f"Evaluator Feedback: {feedback}")

    if "complete and accurate" in feedback.lower():
        print(f"Presentation approved. Saving as {FINAL_OUTPUT_PPT}.")
        shutil.copy(current_output_ppt, FINAL_OUTPUT_PPT)
        break

    iteration += 1

if iteration > MAX_ITERATIONS:
    print("Max iterations reached. Saving the latest version as final.")
    shutil.copy(current_output_ppt, FINAL_OUTPUT_PPT)

print(f"Final presentation saved as {FINAL_OUTPUT_PPT}.")

