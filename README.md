[English](./README.md) | [简体中文](./README.zh-CN.md)

# LLM Document Editor (llm2doc)

This project is a Python-based tool that allows you to edit Microsoft Word (`.docx`) files using natural language commands. It leverages Large Language Models (LLMs) to interpret your instructions and apply complex changes to both the content and styling of a document.

## Features

-   **Comprehensive Content and Style Editing**: Modify a wide range of document properties, including:
    -   **Text-Level Formatting**:
        -   Text content, bold, italic, underline (single, double, etc.).
        -   Font name, size, and color (hex).
        -   Highlight color.
        -   Strikethrough, double-strikethrough, superscript, and subscript.
        -   All caps and small caps.
    -   **Paragraph Formatting**:
        -   Alignment (left, center, right, justify).
        -   Indentation (left, right, and first line).
        -   Spacing (before and after paragraph).
        -   Line spacing (multiple, points, and rules).
        -   Bulleted and numbered lists.
    -   **Structural Elements**:
        -   Create and modify tables.
        -   Add and resolve hyperlinks.
-   **Document Creation**: Generate a complete `.docx` file from a natural language prompt.
-   **Multi-Provider Support**: Works with multiple LLM providers, including OpenAI, Google Gemini, Groq, and others.
-   **Precise Modifications**: Converts the `.docx` to a detailed JSON structure with unique IDs for each element, allowing for highly specific edits.
-   **Command-Line Interface**: A simple CLI (`run.py`) streamlines the entire process with `edit` and `create` commands.
-   **Extensible**: Easily add new LLM providers by implementing a new client.

## How It Works

The tool follows a four-step process:

1.  **Convert to JSON**: The input `.docx` file is converted into a detailed JSON representation that captures its content, structure, and styling.
2.  **LLM Editing**: Your natural language query and the JSON data are sent to the selected LLM. The model analyzes your request and returns a list of modified JSON objects.
3.  **Apply Modifications**: The changes from the LLM are merged back into the original JSON structure.
4.  **Rebuild DOCX**: The updated JSON is used to construct a new `.docx` file with the requested modifications applied.

## Core Components

-   `run.py`: The main command-line interface for running the end-to-end workflow.
-   `llm2doc/converter.py`: Handles the conversion between `.docx` and the project's specific JSON format.
-   `llm2doc/editor.py`: Manages the interaction with the LLM, sending the user's query and processing the response.
-   `llm2doc/processor.py`: Applies the modifications returned by the LLM to the JSON data.
-   `llm2doc/llm_clients.py`: Contains the clients for connecting to various LLM APIs.
-   `llm2doc/prompt/`: This directory contains the system prompts that instruct the LLM on how to handle creation and modification tasks.

## Setup and Usage

### Prerequisites

-   Python 3.7+
-   An API key for the LLM service you intend to use.

### Installation

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd <repository-directory>
    ```

2.  **Install the required packages:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Set Your API Key**
    Set the appropriate environment variable for the LLM provider you plan to use. This is the recommended and most secure method.

    *For Linux/macOS:*
    ```bash
    export OPENAI_API_KEY="your-openai-api-key"
    export GEMINI_API_KEY="your-gemini-api-key"
    export GROQ_API_KEY="your-groq-api-key"
    export DASHSCOPE_API_KEY="your-dashscope-api-key"
    export DEEPSEEK_API_KEY="your-deepseek-api-key"
    export DOUBAO_API_KEY="your-doubao-api-key"
    ```

    *For Windows (Command Prompt):*
    ```bash
    set OPENAI_API_KEY="your-openai-api-key"
    set GEMINI_API_KEY="your-gemini-api-key"
    set GROQ_API_KEY="your-groq-api-key"
    set DASHSCOPE_API_KEY="your-dashscope-api-key"
    set DEEPSEEK_API_KEY="your-deepseek-api-key"
    set DOUBAO_API_KEY="your-doubao-api-key"
    ```

### Running the Tool

The `run.py` script provides two main commands: `edit` and `create`.

#### Editing an Existing Document

To modify an existing `.docx` file, use the `edit` command:

```bash
python run.py edit \
  --input-docx "path/to/your/document.docx" \
  --query "Change the main title to 'Annual Report' and make it bold." \
  --output-docx "path/to/your/modified_document.docx" \
  --provider "chatgpt"
```

#### Creating a New Document

To generate a new `.docx` file from a query, use the `create` command:

```bash
python run.py create \
  --query "Create a document with a title 'My First Document' that is bold and centered. Add a paragraph below it saying 'Hello, world!' in blue text." \
  --output-docx "new_document.docx" \
  --provider "chatgpt"
```

## Supported Providers

You can specify the LLM provider using the `--provider` argument. The supported providers and their corresponding required environment variables are:

| Provider | `--provider` value | Environment Variable  |
|----------|--------------------|-----------------------|
| OpenAI   | `chatgpt`          | `OPENAI_API_KEY`      |
| Google   | `gemini`           | `GEMINI_API_KEY`      |
| Groq     | `groq`             | `GROQ_API_KEY`        |
| Qwen     | `qwen`             | `DASHSCOPE_API_KEY`   |
| Deepseek | `deepseek`         | `DEEPSEEK_API_KEY`    |
| Doubao   | `doubao`           | `DOUBAO_API_KEY`      |

If no provider is specified, it defaults to `chatgpt`.

## How to Use in a Jupyter Notebook

You can run the entire process within a single Jupyter Notebook cell. This is useful for interactive testing and experimentation. Copy and paste the following code into a cell:

```python
# Step 1: Install dependencies if you haven't already
!pip install -r requirements.txt

import os
from docx import Document
from run import create_docx, modify_docx

# --- Configuration ---
# 1. Set your desired LLM provider here
# Supported: "chatgpt", "gemini", "groq", "qwen", "deepseek", "doubao"
PROVIDER = "doubao"

# 2. Set the corresponding API key
# Make sure the environment variable matches the provider you chose.
os.environ["DOUBAO_API_KEY"] = "your-doubao-api-key-here"
# os.environ["OPENAI_API_KEY"] = "your-openai-api-key-here"
# os.environ["GEMINI_API_KEY"] = "your-gemini-api-key-here"


# --- Option 1: Create a new document ---
print(f"--- Running Creation Example with {PROVIDER.upper()} ---")
creation_query = "Create a document with a title 'New Notebook Doc' that is bold. Add a paragraph below it saying 'This was created from a notebook.'"
creation_output = "notebook_created.docx"

try:
    create_docx(
        query=creation_query,
        output_docx=creation_output,
        provider=PROVIDER
    )
    print(f"Successfully created '{creation_output}'")
except Exception as e:
    print(f"An error occurred during creation: {e}")


# --- Option 2: Edit an existing document ---
print(f"\n--- Running Modification Example with {PROVIDER.upper()} ---")
# Create a dummy document for testing
edit_input_file = "notebook_for_editing.docx"
if not os.path.exists(edit_input_file):
    doc = Document()
    doc.add_paragraph("This is the original text. It is black and not bold.")
    doc.save(edit_input_file)
    print(f"Created dummy file: '{edit_input_file}'")

edit_query = "Change the text 'black' to 'purple' and make the font color purple."
edit_output_file = "notebook_modified.docx"

try:
    modify_docx(
        input_docx=edit_input_file,
        query=edit_query,
        output_docx=edit_output_file,
        provider=PROVIDER
    )
    print(f"Successfully modified '{edit_input_file}'. Output saved to '{edit_output_file}'")
except Exception as e:
    print(f"An error occurred during modification: {e}")
```

## Future Work

-   **PPTX Support**: We are planning to extend the functionality of this tool to support editing Microsoft PowerPoint (`.pptx`) files in a similar manner.