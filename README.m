This program is designed to format and clean up text documents in the .docx format. Below is an overview of its functionality:

1. Import necessary libraries:
    - os
    - re
    - PyQt6.QtWidgets
    - sys
    - asyncio
    - docx.Document
    - tqdm

2. Define asynchronous functions for cleaning paragraphs, processing chunks of text, and formatting .docx documents.

3. Split the input document into chunks of paragraphs.

4. Create a new document from the processed chunks.

5. Define a PyQt6-based GUI application for user interaction.

6. Create the GUI layout with input and output file fields and buttons for browsing and formatting.

7. Connect event handlers for browsing input and output files, and formatting the document.

To use the program:
- Run the script.
- Browse and select the input .docx file.
- Browse and specify the output .docx file.
- Click the "Format" button to initiate the formatting process.
- The formatted document will be saved to the specified output location.

Note: The program utilizes asynchronous processing for efficient handling of large documents.
