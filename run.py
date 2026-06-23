
import os
import base64
import importlib
from typing import Any
from openai import AzureOpenAI

dotenv_module = importlib.import_module("dotenv") if importlib.util.find_spec("dotenv") else None
if dotenv_module:
    dotenv_module.load_dotenv()


def require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


endpoint = require_env("ENDPOINT_URL")
deployment = require_env("DEPLOYMENT_NAME")
search_endpoint = require_env("SEARCH_ENDPOINT")
search_key = require_env("SEARCH_KEY")
search_index = require_env("SEARCH_INDEX_NAME")
subscription_key = require_env("OPENAI_API_KEY")

# Document filter configuration
DOCUMENT_ID = "d0e2081f096848789baf8d677d948edb"
search_filter = f"search.ismatch('{DOCUMENT_ID}', 'title')"

# Initialize Azure OpenAI client with key-based authentication
client = AzureOpenAI(
    azure_endpoint=endpoint,
    api_key=subscription_key,
    api_version="2025-01-01-preview",
)

# IMAGE_PATH = "YOUR_IMAGE_PATH"
# encoded_image = base64.b64encode(open(IMAGE_PATH, 'rb').read()).decode('ascii')

# Prepare the chat prompt
chat_prompt = [
    {
        "role": "system",
        "content": "You are a seasoned real estate attorney specializing in the creation and abstraction of lease documents, amendments, and acknowledgment letters. Only respond with information explicitly stated in the provided lease documents. Do not generate, infer, or provide sample content. If information is not found, respond with 'NOT FOUND FOR THIS FIELD'."
    },
    {
        "role": "user",
        "content": "Extract all complete property addresses mentioned in the lease document, including building names/numbers as standalone strings. Include street, city, state, and zip code if available. DO NOT include notice addresses, mailing addresses for rent payment, or any surrounding commentary like 'The property is located at...' or 'Premises known as...'. Only extract the physical property location addresses.",
    }
    
]


# Include speech result if speech is enabled
messages: Any = chat_prompt

# Generate the completion
completion = client.chat.completions.create(
    model=deployment,
    messages=messages,
    max_tokens=6553,
    temperature=0.7,
    top_p=0.95,
    frequency_penalty=0,
    presence_penalty=0,
    stop=None,
    stream=False,
    extra_body={
      "data_sources": [{
          "type": "azure_search",
          "parameters": {
            "endpoint": f"{search_endpoint}",
            "index_name": f"{search_index}",
            "semantic_configuration": "default",
            "query_type": "semantic",
            "fields_mapping": {},
            "in_scope": True,
            "filter": search_filter,
            "strictness": 1,
            "top_n_documents": 10,
            "authentication": {
              "type": "api_key",
              "key": f"{search_key}"
            }
          }
        }]
    }
)

# Extract structured response
response_content = completion.choices[0].message.content
citations = completion.choices[0].message.context.citations if hasattr(completion.choices[0].message, 'context') and hasattr(completion.choices[0].message.context, 'citations') else []

structured_result = {
    "field": "Property Address",  # Replace with dynamic field name when looping
    "response": response_content,
    "citations": citations,
    "tokens_used": completion.usage.completion_tokens,
    "finish_reason": completion.choices[0].finish_reason
}

import json
print(json.dumps(structured_result, indent=2))
    