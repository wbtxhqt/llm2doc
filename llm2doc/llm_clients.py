"""
This module provides a set of clients for interacting with various Large Language Model (LLM) APIs.

Each client is responsible for sending a prompt to a specific LLM service and returning the response.
API keys are expected to be set as environment variables.

Required packages:
- openai
- google-generativeai
- groq
- requests

You can install them using pip:
pip install openai google-generativeai groq requests
"""

import os
import json
from abc import ABC, abstractmethod
from typing import Any, Dict

# It's good practice to use a library like `python-dotenv` to manage environment variables,
# but for simplicity, we'll assume they are pre-loaded.
# from dotenv import load_dotenv
# load_dotenv()

class LLMClient(ABC):
    """Abstract base class for LLM API clients."""

    @abstractmethod
    def invoke(self, system_prompt: str, user_prompt: str, json_data: Dict[str, Any] | None) -> str:
        """
        Sends a request to the LLM and returns the response content.

        Args:
            system_prompt: The system-level instructions for the model.
            user_prompt: The user's specific query.
            json_data: The JSON data to be processed by the LLM, or None for creation.

        Returns:
            The LLM's response as a string.
        """
        pass

class ChatGPTClient(LLMClient):
    """Client for OpenAI's ChatGPT API."""

    def __init__(self, model: str = "gpt-4-turbo"):
        from openai import OpenAI
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("OPENAI_API_KEY environment variable not set.")
        self.client = OpenAI(api_key=api_key)
        self.model = model

    def invoke(self, system_prompt: str, user_prompt: str, json_data: Dict[str, Any] | None) -> str:
        if json_data:
            full_prompt = f"{user_prompt}\n\nHere is the JSON data:\n{json.dumps(json_data, indent=2)}"
        else:
            full_prompt = user_prompt

        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": full_prompt},
            ],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content or ""

class GeminiClient(LLMClient):
    """Client for Google's Gemini API."""

    def __init__(self, model: str = "gemini-1.5-pro-latest"):
        import google.generativeai as genai
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("GEMINI_API_KEY environment variable not set.")
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel(
            model,
            generation_config={"response_mime_type": "application/json"}
        )

    def invoke(self, system_prompt: str, user_prompt: str, json_data: Dict[str, Any] | None) -> str:
        if json_data:
            full_prompt = f"{system_prompt}\n\n{user_prompt}\n\nHere is the JSON data:\n{json.dumps(json_data, indent=2)}"
        else:
            full_prompt = f"{system_prompt}\n\n{user_prompt}"

        response = self.model.generate_content(full_prompt)
        return response.text

class GroqClient(LLMClient):
    """Client for Groq's API."""

    def __init__(self, model: str = "llama3-70b-8192"):
        from groq import Groq
        api_key = os.getenv("GROQ_API_KEY")
        if not api_key:
            raise ValueError("GROQ_API_KEY environment variable not set.")
        self.client = Groq(api_key=api_key)
        self.model = model

    def invoke(self, system_prompt: str, user_prompt: str, json_data: Dict[str, Any] | None) -> str:
        if json_data:
            full_prompt = f"{user_prompt}\n\nHere is the JSON data:\n{json.dumps(json_data, indent=2)}"
        else:
            full_prompt = user_prompt

        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": full_prompt},
            ],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content or ""

class QwenClient(LLMClient):
    """Client for Alibaba's Qwen (Tongyi Qwen) API."""

    def __init__(self, model: str = "qwen-turbo"):
        # Qwen uses an OpenAI-compatible client
        from openai import OpenAI
        api_key = os.getenv("DASHSCOPE_API_KEY") # Alibaba's key
        if not api_key:
            raise ValueError("DASHSCOPE_API_KEY environment variable not set.")
        self.client = OpenAI(
            api_key=api_key,
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"
        )
        self.model = model

    def invoke(self, system_prompt: str, user_prompt: str, json_data: Dict[str, Any] | None) -> str:
        if json_data:
            full_prompt = f"{user_prompt}\n\nHere is the JSON data:\n{json.dumps(json_data, indent=2)}"
        else:
            full_prompt = user_prompt
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": full_prompt},
            ]
        )
        return response.choices[0].message.content or ""


class DoubaoClient(LLMClient):
    """Client for Doubao's API (Volcano Engine)."""

    def __init__(self, model: str = "doubao-seed-1-6-250615"):
        from openai import OpenAI
        api_key = os.getenv("DOUBAO_API_KEY")
        if not api_key:
            raise ValueError("DOUBAO_API_KEY environment variable not set.")
        self.client = OpenAI(
            base_url="https://ark.cn-beijing.volces.com/api/v3",
            api_key=api_key
        )
        self.model = model

    def invoke(self, system_prompt: str, user_prompt: str, json_data: Dict[str, Any] | None) -> str:
        if json_data:
            full_prompt = f"{user_prompt}\n\nHere is the JSON data:\n{json.dumps(json_data, indent=2)}"
        else:
            full_prompt = user_prompt
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": full_prompt},
            ]
        )
        return response.choices[0].message.content or ""

class DeepseekClient(LLMClient):
    """Client for Deepseek's API."""

    def __init__(self, model: str = "deepseek-chat"):
        # Deepseek also uses an OpenAI-compatible client
        from openai import OpenAI
        api_key = os.getenv("DEEPSEEK_API_KEY")
        if not api_key:
            raise ValueError("DEEPSEEK_API_KEY environment variable not set.")
        self.client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com/v1")
        self.model = model

    def invoke(self, system_prompt: str, user_prompt: str, json_data: Dict[str, Any] | None) -> str:
        if json_data:
            full_prompt = f"{user_prompt}\n\nHere is the JSON data:\n{json.dumps(json_data, indent=2)}"
        else:
            full_prompt = user_prompt
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": full_prompt},
            ],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content or ""

def get_llm_client(provider: str) -> LLMClient:
    """
    Factory function to get an LLM client based on the provider name.

    Args:
        provider: The name of the LLM provider. 
                  Supported: "chatgpt", "gemini", "groq", "qwen", "deepseek", "doubao".

    Returns:
        An instance of the corresponding LLMClient.
    """
    provider = provider.lower()
    if provider == "chatgpt":
        return ChatGPTClient()
    elif provider == "gemini":
        return GeminiClient()
    elif provider == "groq":
        return GroqClient()
    elif provider == "qwen":
        return QwenClient()
    elif provider == "deepseek":
        return DeepseekClient()
    elif provider == "doubao":
        return DoubaoClient()
    else:
        raise ValueError(f"Unknown LLM provider: {provider}")

# Example usage:
if __name__ == "__main__":
    # Make sure to set the respective API key in your environment,
    # e.g., export OPENAI_API_KEY="your-key"
    
    # --- Mock Data ---
    mock_system_prompt = "You are a helpful assistant. The user will provide a JSON object. Modify it as requested and return the modified JSON."
    mock_user_prompt = "Change the name to 'Jane Doe' and the age to 31."
    mock_json = {"name": "John Doe", "age": 30, "city": "New York"}

    # --- Select Provider ---
    # provider_name = "chatgpt" 
    # provider_name = "gemini"
    # provider_name = "groq"
    provider_name = "qwen" # Change this to test different providers

    try:
        print(f"--- Testing {provider_name.upper()} Client ---")
        client = get_llm_client(provider_name)
        
        # Check for API key before invoking
        key_map = {
            "chatgpt": "OPENAI_API_KEY",
            "gemini": "GEMINI_API_KEY",
            "groq": "GROQ_API_KEY",
            "qwen": "DASHSCOPE_API_KEY",
            "deepseek": "DEEPSEEK_API_KEY"
        }
        if not os.getenv(key_map.get(provider_name)):
             print(f"Skipping test: {key_map.get(provider_name)} is not set.")
        else:
            response_text = client.invoke(mock_system_prompt, mock_user_prompt, mock_json)
            print("LLM Response:")
            print(response_text)

    except (ValueError, ImportError) as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")