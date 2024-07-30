# Thematic Analysis Script

This script performs a thematic analysis of a Google Document using OpenAI's GPT-4 model. It generates initial codes and descriptions based on the document's content and metadata.
Permission: Ulysses Cabayao, SJ 2024 (uscabayaosj@addu.edu.ph)

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
- [How it Works](#how-it-works)
- [Functions](#functions)
- [Security](#security)
- [Troubleshooting](#troubleshooting)

## Installation
1. Create a new Google Apps Script project.
2. Copy and paste the script into the editor.
3. Set up the OpenAI API key in the Script Properties:
   - Go to the Script Properties page (File > Properties).
   - Add a new property named `OPENAI_API_KEY` and enter your API key.
4. Save the script.

## Usage
1. Open a Google Document.
2. Click on the "Thematic Analysis" menu and select "Generate Initial Codes".
3. The script will create a new document with the initial codes and descriptions.
4. A link to the new document will be appended to the original document.

## How it Works
The script uses OpenAI's GPT-4 model to perform a thematic analysis of the document's content and metadata. It follows Braun and Clarke's six-phase guide, focusing on steps 1 and 2 to generate initial codes. The script then creates a new document with the codes and descriptions in a tabular format.

## Functions
- `onOpen()`: Creates the "Thematic Analysis" menu.
- `analyzeDocument()`: Performs the thematic analysis and generates the initial codes.
- `createCodesTable()`: Creates a table for the codes and descriptions.
- `getOpenAIInsight()`: Fetches the insight from OpenAI's API.
- `createOverallAssessment()`: Creates a new document with the initial codes and descriptions.
- `appendAssessmentLink()`: Appends a link to the new document to the original document.
- `getApiKey()`: Retrieves the OpenAI API key from the Script Properties.

## Security
The script uses the `PropertiesService` to store the OpenAI API key securely. Make sure to keep your API key confidential and do not share it with anyone.

## Troubleshooting
- Check the logs for any errors.
- Make sure the OpenAI API key is set up correctly.
- Ensure that the document has the necessary permissions for the script to run.

**Note**: This script is for educational purposes only and should not be used for production without proper testing and validation.
