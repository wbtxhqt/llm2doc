import argparse
import json
import os
import tempfile
from typing import Any, Dict, List

import config
from llm2doc.converter import docx2json, json2docx
from llm2doc.editor import create_document, edit_document
from llm2doc.processor import apply_modifications


def create_docx(query: str, output_docx: str, provider: str = config.DEFAULT_PROVIDER):
    """
    Creates a new .docx file based on a user query using an LLM.

    Args:
        query: The natural language instruction for creating the document.
        output_docx: Path to save the new .docx file.
        provider: The LLM provider to use.
    """
    print("Starting creation workflow...")
    print("\nStep 1: Sending request to LLM for document creation...")
    try:
        created_json = create_document(user_query=query, provider=provider)
        if not created_json:
            print("LLM failed to generate a document. Exiting.")
            return
        print("Successfully received created document from LLM.")
    except Exception as e:
        print(f"Error during LLM creation step: {e}")
        return

    print(f"\nStep 2: Converting created JSON to '{output_docx}'...")
    try:
        json2docx(created_json, output_docx)
        print(f"Successfully created new .docx file: '{output_docx}'")
    except Exception as e:
        print(f"Error during JSON to DOCX conversion: {e}")


def modify_docx(input_docx: str, query: str, output_docx: str, provider: str = config.DEFAULT_PROVIDER, keep_temp_file: bool = False):
    """
    Modifies a .docx file based on a user query using an LLM.

    Args:
        input_docx: Path to the input .docx file.
        query: The natural language instruction for the edit.
        output_docx: Path to save the modified .docx file.
        provider: The LLM provider to use.
        keep_temp_file: If True, the intermediate JSON file will not be deleted.
    """
    temp_json_file = tempfile.NamedTemporaryFile(
        mode='w+', delete=False, suffix=".json", encoding="utf-8"
    )
    temp_json_path = temp_json_file.name
    temp_json_file.close()

    try:
        print(f"Step 1: Converting '{input_docx}' to JSON...")
        try:
            full_json_data = docx2json(input_docx, json_path=temp_json_path, compact=False)
            print(f"Successfully converted to '{temp_json_path}'")
        except Exception as e:
            print(f"Error during DOCX to JSON conversion: {e}")
            return

        print("\nStep 2: Sending request to LLM for editing...")
        try:
            modified_parts = edit_document(
                user_query=query,
                json_path=temp_json_path,
                provider=provider
            )
            if not modified_parts:
                print("LLM returned no modifications. Exiting.")
                return
            print("Successfully received modifications from LLM.")
            print(json.dumps(modified_parts, indent=2))
        except Exception as e:
            print(f"Error during LLM editing step: {e}")
            return

        print("\nStep 3: Applying modifications to the full JSON data...")
        updated_json_data = apply_modifications(full_json_data, modified_parts)

        print(f"\nStep 4: Converting modified JSON back to '{output_docx}'...")
        try:
            json2docx(updated_json_data, output_docx)
            print("Successfully created modified .docx file.")
        except Exception as e:
            print(f"Error during JSON to DOCX conversion: {e}")

    finally:
        if not keep_temp_file and os.path.exists(temp_json_path):
            os.remove(temp_json_path)
            print(f"Cleaned up temporary file: '{temp_json_path}'")
        elif keep_temp_file:
            print(f"Temporary JSON file saved at: '{temp_json_path}'")


def main():
    """
    Main function to handle command-line operations for docx modification and creation.
    """
    parser = argparse.ArgumentParser(description="Create or edit a .docx file using an LLM.")
    subparsers = parser.add_subparsers(dest="command", required=True, help="Available commands")

    # --- Edit Command ---
    parser_edit = subparsers.add_parser("edit", help="Edit an existing .docx file.")
    parser_edit.add_argument("--input-docx", required=True, help="Path to the input .docx file.")
    parser_edit.add_argument("--query", required=True, help="The natural language instruction for the edit.")
    parser_edit.add_argument("--output-docx", required=True, help="Path to save the modified .docx file.")
    parser_edit.add_argument("--provider", default=config.DEFAULT_PROVIDER, help="The LLM provider to use.")
    parser_edit.add_argument("--keep-temp-file", action="store_true", help="Keep the intermediate JSON file.")

    # --- Create Command ---
    parser_create = subparsers.add_parser("create", help="Create a new .docx file from a query.")
    parser_create.add_argument("--query", required=True, help="The natural language instruction for creating the document.")
    parser_create.add_argument("--output-docx", required=True, help="Path to save the new .docx file.")
    parser_create.add_argument("--provider", default=config.DEFAULT_PROVIDER, help="The LLM provider to use.")

    args = parser.parse_args()

    if args.command == "edit":
        modify_docx(
            input_docx=args.input_docx,
            query=args.query,
            output_docx=args.output_docx,
            provider=args.provider,
            keep_temp_file=args.keep_temp_file
        )
    elif args.command == "create":
        create_docx(
            query=args.query,
            output_docx=args.output_docx,
            provider=args.provider
        )


if __name__ == "__main__":
    main()