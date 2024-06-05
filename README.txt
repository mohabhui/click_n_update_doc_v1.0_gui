**Application Name:** click-n-update  
**Application Version:** 1.0  
**Author:** mohabhui  
**Date:** 05-Jun-2024  

#### Description

click-n-update is a Python GUI application designed to update text surrounded by double curly brackets (placeholders), e.g., `{{applicant_name}}`, in a Word document using key-value pairs from a JSON file. The application ensures that the specified text is replaced and formatted according to the JSON file's instructions.

#### Sample JSON File (`sample.json`):

```json
{
  "template_doc": ["sample.docx"],
  "applicant_name": ["John Doe", "Arial", "Bold", "Italic", "Blue", "Underline"],
  "job_title": ["Software Engineer", "Cyan", "Courier New", "Bold"],
  "joining_date": ["01-Sep-2024", "Red"],
  "hourly_rate": ["CA$ 85", "Bold", "Purple"]
}
```

#### Sample Word Document (`sample.docx`):

```
Dear {{applicant_name}},
We are pleased to offer you the position of {{job_title}} at our company.
Applicant Name: {{applicant_name}}
Job Title: {{job_title}}
Joining Date: {{joining_date}}
Hourly Rate: {{hourly_rate}}
Best regards,
______________________________
Hiring Team
```

#### Key Features:

- **Placeholder Replacement:** The text surrounded by double curly brackets in the Word document are keys of the JSON file. The application replaces these placeholders with corresponding values from the JSON file.
- **Formatting Application:** The application applies formatting specified in the JSON file to the replaced text. Allowed formatting includes:
  - **Font Name:** e.g., "Arial"
  - **Font Type:** "Bold", "Italic", "Underline"
  - **Font Color:** "Red", "Green", "Blue", "Purple", "Cyan", "Orange", "Yellow", "Black", "White"
- **GUI Functionality:**
  - **Browse Button:** Allows users to browse and select the JSON file.
  - **Dynamic Display:** Displays the JSON file's key-value pairs dynamically in the GUI.
  - **Save Button:** Initially inactive and becomes active when:
    1. The JSON file is opened.
    2. The JSON file contains the key-value for the document.
    3. There is at least one key-value for the text to be replaced.
- **Multiple Occurrences:** Supports replacement of the same key in multiple locations within the Word document.
- **Formatting Preservation:** Preserves existing formatting of the Word document while applying new formatting to the updated text.

#### Usage:

1. **Open the Application:** Run the Python script to open the GUI.
2. **Browse JSON File:** Use the "Browse JSON File" button to select the JSON file containing the key-value pairs.
3. **View JSON Content:** The key-value pairs from the JSON file will be displayed dynamically in the GUI.
4. **Activate Save Button:** Ensure that the JSON file meets the conditions to activate the "Save" button.
5. **Update Document:** Press the "Save" button to update the Word document with the specified text and formatting, then save the updated document.

This application ensures that placeholders in the Word document are replaced with the specified text from the JSON file and formatted accordingly, providing an efficient and user-friendly solution for document updates.