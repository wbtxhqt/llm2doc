import json
from typing import Any, Dict, List
from .llm_clients import get_llm_client


def create_document(
    user_query: str,
    provider: str,
    prompt_path: str = "llm2doc/prompt/create_doc_prompt.txt"
) -> Dict[str, Any]:
    """
    Sends a user query to an LLM to generate a new document in JSON format.

    Args:
        user_query: The user's instruction for creating the document.
        provider: The LLM provider to use.
        prompt_path: The path to the system prompt for creation.

    Returns:
        A dictionary representing the complete JSON of the new document.
    """
    with open(prompt_path, "r", encoding="utf-8") as f:
        system_prompt = f.read()

    print(f"--- Sending request to {provider.upper()} for document creation ---")
    
    try:
        llm_client = get_llm_client(provider)
        # For creation, we don't provide existing JSON data, only the query
        created_json_str = llm_client.invoke(system_prompt, user_query, json_data=None)
        
        created_json = json.loads(created_json_str)
        
        if not isinstance(created_json, dict):
            print("Warning: LLM did not return a valid JSON object for the document.")
            return {}
            
        return created_json

    except (ValueError, ImportError) as e:
        print(f"Error initializing LLM client: {e}")
        return {}
    except json.JSONDecodeError:
        print("Error: The LLM did not return valid JSON.")
        return {}
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return {}


def edit_document(
    user_query: str,
    json_path: str,
    provider: str,
    prompt_path: str = "llm2doc/prompt/modify_doc_prompt.txt"
) -> List[Dict[str, Any]]:
    """
    Loads a document in JSON format, combines it with a user query and a prompt,
    sends it to the specified LLM provider, and returns the modified objects.

    Args:
        user_query: The user's instruction for modifying the document.
        json_path: The path to the JSON file representing the document.
        provider: The LLM provider to use (e.g., "chatgpt", "gemini").
        prompt_path: The path to the system prompt file.

    Returns:
        A list of dictionaries, where each dictionary is a modified object.
    """
    # Load the system prompt
    with open(prompt_path, "r", encoding="utf-8") as f:
        system_prompt = f.read()

    # Load the document from the specified JSON file
    with open(json_path, "r", encoding="utf-8") as f:
        json_data = json.load(f)

    print(f"--- Sending request to {provider.upper()} ---")
    
    try:
        # Get the client and invoke the LLM
        llm_client = get_llm_client(provider)
        modified_json_str = llm_client.invoke(system_prompt, user_query, json_data)
        
        # The LLM is expected to return a JSON array of modified objects
        modified_objects = json.loads(modified_json_str)
        
        if not isinstance(modified_objects, list):
            print("Warning: LLM did not return a JSON array. Wrapping in a list.")
            return [modified_objects]
            
        return modified_objects

    except (ValueError, ImportError) as e:
        print(f"Error initializing LLM client: {e}")
        return []
    except json.JSONDecodeError:
        print("Error: The LLM did not return valid JSON.")
        return []
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return []

if __name__ == "__main__":
    # --- Configuration ---
    # Choose your LLM provider here: "chatgpt", "gemini", "groq", "qwen", "deepseek"
    LLM_PROVIDER = "chatgpt"
    
    TEST_QUERY = "In the paragraph with id 'doc-obj-26', find the run with text 'INSERT Reseller/Distributor Entity Name & Address' and change its text to 'My Company Inc. & 123 Main St'."
    TEST_JSON_PATH = "output.json" # Using the non-compact version with IDs
    
    print(f"Provider: {LLM_PROVIDER.upper()}")
    print(f"Editing '{TEST_JSON_PATH}' with query: '{TEST_QUERY}'")
    
    # Run the edit function
    modified_parts = edit_document(TEST_QUERY, TEST_JSON_PATH, LLM_PROVIDER)
    
    if modified_parts:
        print("\n--- LLM returned modified objects ---")
        print(json.dumps(modified_parts, indent=2))
        
        # (Optional) To apply these changes back to the original file:
        # 1. Load the original JSON
        with open(TEST_JSON_PATH, "r", encoding="utf-8") as f:
            original_data = json.load(f)
        
        # 2. Create a map of objects by ID for efficient lookup
        id_map = {}
        def build_id_map(obj):
            if isinstance(obj, dict) and 'id' in obj:
                id_map[obj['id']] = obj
            if isinstance(obj, dict):
                for value in obj.values():
                    build_id_map(value)
            elif isinstance(obj, list):
                for item in obj:
                    build_id_map(item)
        
        build_id_map(original_data)

        # 3. Update the objects in the map
        for modified_obj in modified_parts:
            if 'id' in modified_obj and modified_obj['id'] in id_map:
                id_map[modified_obj['id']].update(modified_obj)
                print(f"Updated object with ID: {modified_obj['id']}")

        # 4. Save the fully updated document
        output_path = "output_modified.json"
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(original_data, f, indent=2)
        
        print(f"\nFully merged and modified JSON saved to '{output_path}'")
    else:
        print("\n--- No modifications were returned by the LLM ---")