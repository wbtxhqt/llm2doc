from typing import Any, Dict, List

def apply_modifications(original_data: Dict[str, Any], modified_parts: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Applies modifications returned by the LLM to the original JSON data.

    Args:
        original_data: The full JSON data of the document.
        modified_parts: A list of modified objects returned by the LLM.

    Returns:
        The fully updated JSON data.
    """
    # Create a map of all objects by ID for efficient lookup
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

    # Update the objects in the map based on the modifications
    for modified_obj in modified_parts:
        if 'id' in modified_obj and modified_obj['id'] in id_map:
            # Update the existing object with the new values
            id_map[modified_obj['id']].update(modified_obj)
            print(f"Applied update to object with ID: {modified_obj['id']}")
        else:
            print(f"Warning: Could not find object with ID: {modified_obj.get('id')}")
            
    return original_data