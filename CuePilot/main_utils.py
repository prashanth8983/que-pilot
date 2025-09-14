from CuePilot.prompts import *
from CuePilot.model import ModelInterface
from string import Template

import json

def parse_json_response(response: str):
    """
    Extracts and parses a JSON object or array from a string.

    It finds the first occurrence of '{' or '[' and the last occurrence
    of '}' or ']' to identify and parse the JSON substring.
    """
    try:
        # Find the index of the first opening bracket ('{' or '[')
        # We need to handle the case where one of them is not found (-1)
        start_brace = response.find('{')
        start_bracket = response.find('[')
        if start_brace == -1:
            start_index = start_bracket
        elif start_bracket == -1:
            start_index = start_brace
        else:
            start_index = min(start_brace, start_bracket)
        # If no opening bracket is found, there's no JSON to parse
        if start_index == -1:
            # print("No JSON object or array found in the response.")
            return None
        # Find the index of the last closing bracket ('}' or ']')
        end_brace = response.rfind('}')
        end_bracket = response.rfind(']')
        end_index = max(end_brace, end_bracket)
        # If no closing bracket is found, or it's before the opening one, it's invalid
        if end_index == -1 or end_index < start_index:
            # print("Malformed JSON structure: couldn't find a valid end bracket.")
            return None
        # Extract the potential JSON substring
        json_str = response[start_index : end_index + 1]
        # Attempt to parse the extracted string
        return json.loads(json_str)

    except json.JSONDecodeError as e:
        print(f"Failed to parse extracted JSON substring: {e}")
        # print(f"Substring attempted: {json_str}") # Uncomment for debugging
        return None

class PresentationContentManagement:
    def __init__(self):
        pass

    def add_content(self,content:dict):
        self.ppt_flow=content
    
    def get_content(self):
        return self.ppt_flow
    
    def get_content_slide(self,slide_num:str=""):
        return self.ppt_flow.get(slide_num,"")

class PresentationFlow:
    def __init__(self):
        self.model=ModelInterface(model="MODEL_GEMMA34BE")

    def generate_flow(self,script:str,slides_visual_description:str):
        messages=Template(PredefinedPrompts.ppt_flow_generator).\
            safe_substitute(SCRIPT=script,
                            VISUAL_DESCRIPTION=slides_visual_description)
        print("Sending to LLM...")
        response=self.model.chat_completion(messages=[{"role":"user","content":messages}])
        print("Received response from LLM.")
        parsed_response=parse_json_response(response)
        self.ppt_flow=PresentationContentManagement()
        self.ppt_flow.add_content(parsed_response)
        return parsed_response
    
    def getPPTCMS(self):
        return self.ppt_flow        
        