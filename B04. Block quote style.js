const paragraphStyle_EXTRACTED_QUOTE = {
  namedStyleType: 'NORMAL_TEXT',
  indentStart: { magnitude: 35.88159075, unit: 'PT' },
  indentEnd: { magnitude: 35.88159075, unit: 'PT' },
  indentFirstLine: { magnitude: 35.88159075, unit: 'PT' },
  spaceAbove: { magnitude: 20, unit: 'PT' },
  spaceBelow: { magnitude: 20, unit: 'PT' },
  spacingMode: 'NEVER_COLLAPSE',
  alignment: 'START'
};

const textStyle_EXTRACTED_QUOTE_1 = {
  fontSize: {
    magnitude: 11,
    unit: 'PT'
  },
  italic: true,
  weightedFontFamily: {
    fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
    weight: 400
  }
};

const textStyle_EXTRACTED_QUOTE_2 = {
  fontSize: {
    magnitude: 11,
    unit: 'PT'
  },
  italic: false,
  weightedFontFamily: {
    fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
    weight: 400
  }
};

function insertExtractedQuote() {
  const ui = DocumentApp.getUi();
  try {
    const doc = DocumentApp.getActiveDocument();
    const documentId = doc.getId();

    const cursorPosition = detectCursorPosition(doc, documentId);
    if (cursorPosition.status == 'error') {
      ui.alert(cursorPosition.message);
      return 0;
    }

    const insertIndex = cursorPosition.endIndex;
    const requests = [];

    const text = '“Extracted quote goes here” (Author, 2021).';

    requests.push(
      {
        insertText: {
          location: {
            index: insertIndex
          },
          text: text
        }
      },
      {
        updateParagraphStyle: {
          paragraphStyle: paragraphStyle_EXTRACTED_QUOTE,
          range: {
            startIndex: insertIndex,
            endIndex: insertIndex + text.length
          },
          fields: formFieldsString(paragraphStyle_EXTRACTED_QUOTE)
        }
      },
      {
        updateTextStyle: {
          range: {
            startIndex: insertIndex,
            endIndex: insertIndex + 27
          },
          text_style: textStyle_EXTRACTED_QUOTE_1,
          fields: formFieldsString(textStyle_EXTRACTED_QUOTE_1)
        }
      },
      {
        updateTextStyle: {
          range: {
            startIndex: insertIndex + 28,
            endIndex: insertIndex + text.length
          },
          text_style: textStyle_EXTRACTED_QUOTE_2,
          fields: formFieldsString(textStyle_EXTRACTED_QUOTE_2)
        }
      },
      {
        deleteNamedRange: {
          name: cursorPosition.rangeName
        }
      },
      {
        deleteContentRange: {
          range: {
            startIndex: insertIndex - 1,
            endIndex: insertIndex,
          }
        }
      });

    Docs.Documents.batchUpdate({
      requests: requests
    }, documentId);
  }
  catch (error) {
    ui.alert('Error in insertExtractedQuote: ' + error);
    return 0;
  }
}