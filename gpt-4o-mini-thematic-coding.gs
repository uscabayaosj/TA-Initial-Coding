/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Thematic Analysis')
    .addItem('Generate Initial Codes', 'analyzeDocument')
    .addToUi();
}

function analyzeDocument() {
  const doc = DocumentApp.getActiveDocument();
  const file = DriveApp.getFileById(doc.getId());

  // Get document metadata
  const metadata = {
    name: file.getName(),
    dateCreated: file.getDateCreated(),
    dateUpdated: file.getLastUpdated(),
    owner: file.getOwner().getName(),
    editors: file.getEditors().map(editor => editor.getName()).join(', '),
    content: doc.getBody().getText()
  };

  // Create overall assessment and get the URL
  const assessmentUrl = createOverallAssessment(metadata);

  // Append link to the assessment in the original document
  appendAssessmentLink(doc, assessmentUrl);
}

function createCodesTable(body, insight) {
  // Parse the insight to extract codes and descriptions
  const lines = insight.split('\n');
  const codesAndDescriptions = lines.filter(line => line.trim().startsWith('|') && !line.includes('Code') && !line.includes('---'));

  if (codesAndDescriptions.length > 0) {
    const table = body.appendTable();
    const headerRow = table.appendTableRow();
    headerRow.appendTableCell('Code').setBackgroundColor('#f1f3f4');
    headerRow.appendTableCell('Description').setBackgroundColor('#f1f3f4');

    codesAndDescriptions.forEach(line => {
      const [_, code, description] = line.split('|').map(item => item.trim());
      const row = table.appendTableRow();
      row.appendTableCell(code);
      row.appendTableCell(description);
    });

    // Style the table
    table.setColumnWidth(0, 150);
    table.setColumnWidth(1, 350);
    table.setBorderWidth(0);
  } else {
    body.appendParagraph('No codes were generated from the analysis.');
  }
}

function getOpenAIInsight(metadata) {
  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    'model': 'gpt-4o-mini',
    'messages': [
      {'role': 'system', 'content': 'Perform a thematic analysis of the following document using Braun and Clarke\'s six-phase guide. Go through only steps 1 and 2 as part of your thinking process. Focus only on generating initial codes. Present the initial codes in a tabular format with two columns: *Code* and *Description*.'},
      {'role': 'user', 'content': `Document Name: ${metadata.name}\nCreated: ${metadata.dateCreated}\nLast Updated: ${metadata.dateUpdated}\nOwner: ${metadata.owner}\nEditors: ${metadata.editors}\n\nContent: ${metadata.content.substring(0, 1000)}...`}
    ]
  };

  const options = {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + getApiKey(),  // Fetch API key securely
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const jsonResponse = JSON.parse(response.getContentText());
    return jsonResponse.choices[0].message.content;
  } catch (error) {
    Logger.log('Error fetching insights from OpenAI: ' + error);
    return 'Error fetching insights. Please check the logs for more details.';
  }
}

function createOverallAssessment(metadata) {
  const overallInsight = getOpenAIInsight(metadata);
  
  const doc = DocumentApp.create('Initial Coding - ' + metadata.name);
  const body = doc.getBody();
  
  body.appendParagraph('Initial Codes and Descriptions \n')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  // Create a table for codes and descriptions in the overall assessment
  createCodesTable(body, overallInsight);
  
  doc.saveAndClose();
  
  return doc.getUrl();
}

function appendAssessmentLink(originalDoc, assessmentUrl) {
  const body = originalDoc.getBody();
  body.appendParagraph('\n\nDocument Coding Complete')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const linkText = body.appendParagraph('Click here to view the initial codes and their descriptions.');
  linkText.setLinkUrl(assessmentUrl);
}

// Function to get API key from Script Properties
function getApiKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('OPENAI_API_KEY');
}
