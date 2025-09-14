

import json

class ScriptChunker:
    """
    A class to chunk presentation scripts based on slides and add summaries.
    """

    def chunk_script(self, script_text: str) -> str:
        """
        Chunks a presentation script by slides, prepends a summary to each chunk,
        and returns a JSON mapping of slide numbers to the summarized content.

        Args:
            script_text: The presentation script as a single string.
                         Slides are expected to be separated by "\n--- slide ---\n".

        Returns:
            A JSON string mapping slide numbers to the summarized content.
        """
        slides = script_text.strip().split('\n--- slide ---\n')
        slide_map = {}

        for i, slide_content in enumerate(slides):
            slide_number = i + 1
            summary = self._generate_summary(slide_content)
            summarized_content = f"{summary}\n\n{slide_content}"
            slide_map[slide_number] = summarized_content

        return json.dumps(slide_map, indent=4)

    def _generate_summary(self, text: str) -> str:
        """
        Generates a short summary for a given text.

        NOTE: This is a placeholder implementation. A real implementation would
        use a language model to generate a more meaningful summary.

        Args:
            text: The text to summarize.

        Returns:
            A short summary of the text.
        """
        # Placeholder: Take the first sentence as a summary.
        first_sentence = text.split('.')[0]
        if len(first_sentence.split()) > 15: # If the first sentence is too long
            return " ".join(first_sentence.split()[:15]) + "..."
        return first_sentence + "."

if __name__ == '__main__':
    # Example Usage
    script = """This is the first slide. It's about the introduction of our topic. We will discuss the agenda for today.

--- slide ---

This is the second slide. Here we will dive deep into the technical details. We will cover the architecture and the data flow.

--- slide ---

This is the final slide. We will conclude our presentation and have a Q&A session. Thank you for your attention."""

    chunker = ScriptChunker()
    chunked_script_json = chunker.chunk_script(script)
    print(chunked_script_json)

