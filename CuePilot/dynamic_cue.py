from CuePilot.model import ModelInterface
from CuePilot.prompts import PredefinedPrompts
from string import Template
import json
from CuePilot.main_utils import parse_json_response, PresentationContentManagement

testtest="""
{'1': {'Description': "Introduce the topic of project management and its importance in achieving business goals. This slide sets the stage for the presentation's focus on effective project management.", 'slide_rep': 'Project Management: Driving Business Success', 'script': "Good morning, everyone. Today, we'll be discussing the crucial role of project management in achieving business success. Effective project management isn't just about timelines and budgets; it's about strategically guiding initiatives to deliver tangible results."},
"""

class DynamicCue:
    def __init__(self):
        self.model=ModelInterface("MODEL_GEMMA34BE")

    def generate_dynamic_cue(self,user_script:str,slide_num:str,cms=None):
        
        #messages=Template(PredefinedPrompts.dynamic_cue_generator).\
        #    safe_substitute(USER_TEXT=user_script, FULL_SCRIPT=cms.get_content_slide(slide_num=str(slide_num)))
        messages=Template(PredefinedPrompts.dynamic_cue_generator).\
            safe_substitute(USER_TEXT=user_script, FULL_SCRIPT=testtest)
        response=self.model.chat_completion(messages=[{"role":"user","content":messages}])
        print(f"Message: {messages} Second Response by LLM: {response}")
        parsed_response=parse_json_response(response)
        return parsed_response