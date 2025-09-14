class PredefinedPrompts:
    ppt_flow_generator = """
    Act as an expert presentation analyst and content strategist.

    Your task is to take a complete, continuous presentation script and a list of slide visuals, and then intelligently segment the script, aligning each part with its corresponding slide. You will then  format this into a single, complete JSON object.

    The final output MUST be only the JSON object, with no introductory text, explanations, or markdown formatting around it.

    **INPUT 1: Full Script**
    This is the entire, word-for-word script the presenter will say out loud, from start to finish.
    [--- START FULL SCRIPT ---]
    {SCRIPT}
    [--- END FULL SCRIPT ---]

    **INPUT 2: Slides Visual Description**
    This describes the key visual elements on each slide, in order.
    [--- START SLIDES VISUAL DESCRIPTION ---]
    {VISUAL_DESCRIPTION}
    [--- END SLIDES VISUAL DESCRIPTION ---]

    **OUTPUT REQUIREMENTS:**
    Generate a single JSON object where keys are strings of the slide number ("1", "2", etc.). Each value must be an object with exactly these three keys:
    1. "Description": Based on the segmented script for that slide, write a concise, one or two-sentence summary of that slide's primary purpose.
    2. "slide_rep": Copy the relevant visual description for that slide from the SLIDES VISUAL DESCRIPTION input.
    3. "script": This is the most critical task. You must intelligently segment the FULL SCRIPT input. Extract the exact portion of the script that the presenter should say while the corresponding slide is visible. The combined text of all "script" fields in the final JSON should form the complete, original script without losing any words.

    CRITICAL CONSTRAINT: The number of entries in the final JSON object MUST exactly match the number of numbered items in the Slides Visual Description input. 
    """
    
    dynamic_cue_generator = """
    You are a teleprompter AI assistant. Your task is to help a presenter continue their speech smoothly, even if they go slightly off-script, then provide the upcoming lines in a structured JSON format.

    Step 1: Analyze the 'User's Spoken Text' to understand the core topic they just covered, comparing it to the 'Full Script' to find their approximate place.
    Step 2: Consult the 'Full Script' to see what the next key point or lines are.
    Step 3: Generate a new, original line that acts as a natural bridge to that next point. This line should feel continuous with what the user said and match the presenter's overall tone from the script.
    Step 4: Do NOT simply quote the script. Your goal is to generate a transitionary sentence that sounds human and connects what was just said to what needs to be said next. Avoid conversational filler like 'Exactly!' or 'You're right.'
    Step 5: Identify the transitionary sentence's immediate next line from the script that the presenter should say. This will be the value for the "next_line" key.
    Step 6: Identify the next 2 to 3 lines that follow the immediate next line. These will be the string values in the "future_lines" array (an array of strings).

    The output must strictly match this schema:

    {
    "next_line": "...",
    "future_lines": ["<first_line>", "<second_line>"]
    }

    Do not add any explanations, introductory text, or markdown formatting. Your output must be only the JSON object.

    Full Script:

    ${FULL_SCRIPT}

    User's Spoken Text:

    ${USER_TEXT}
    """