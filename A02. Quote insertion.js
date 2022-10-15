const zeroMagnitude = {
  magnitude: 0,
  unit: 'PT'
};

const paragraphStyle_QUOTE_1 = {
  namedStyleType: 'NORMAL_TEXT',
  indentStart: { magnitude: 56.692913385826756, unit: 'PT' },
  indentEnd: { magnitude: 75.75, unit: 'PT' },
  indentFirstLine: { magnitude: 56.692913385826756, unit: 'PT' },
  spaceAbove: { magnitude: 18, unit: 'PT' },
  spaceBelow: { magnitude: 10, unit: 'PT' },
  spacingMode: 'NEVER_COLLAPSE',
  alignment: 'START',
  borderTop: {
    width: {
      magnitude: 1.5,
      unit: 'PT'
    },
    padding: {
      magnitude: 4,
      unit: 'PT'
    },
    dashStyle: 'SOLID',
    color: {
      color: {
        rgbColor: hexToRGB(styles[ACTIVE_STYLE]['main_heading_font_color'])
      }
    },
  }
};

const paragraphStyle_QUOTE_2 = {
  namedStyleType: 'NORMAL_TEXT',
  indentStart: { magnitude: 56.692913385826756, unit: 'PT' },
  indentEnd: { magnitude: 75.75, unit: 'PT' },
  indentFirstLine: { magnitude: 56.692913385826756, unit: 'PT' },
  spaceAbove: { magnitude: 10, unit: 'PT' },
  spaceBelow: { magnitude: 18, unit: 'PT' },
  spacingMode: 'NEVER_COLLAPSE',
  alignment: 'END',
  borderBottom: {
    width: {
      magnitude: 1.5,
      unit: 'PT'
    },
    padding: {
      magnitude: 4,
      unit: 'PT'
    },
    dashStyle: 'SOLID',
    color: {
      color: {
        rgbColor: hexToRGB(styles[ACTIVE_STYLE]['main_heading_font_color'])
      }
    },
  }
};

const textStyle_QUOTE = {
  foregroundColor: {
    color: {
      rgbColor: hexToRGB(styles[ACTIVE_STYLE]['main_heading_font_color'])
    }
  },
  fontSize: {
    magnitude: 14,
    unit: 'PT'
  },
  bold: true,
  weightedFontFamily: {
    fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
    weight: 400
  }
};

// insertPullQuote use the function
function addQuoteAPIrequest(requests, text, insertIndex, paragraphStyle, textStyle) {
  requests.push(
    {
      insertText: {
        location: {
          index: insertIndex
        },
        text: text
      }
    }, {
    updateParagraphStyle: {
      paragraphStyle: paragraphStyle,
      range: {
        startIndex: insertIndex,
        endIndex: insertIndex + text.length
      },
      fields: formFieldsString(paragraphStyle)
    }
  }, {
    updateTextStyle: {
      range: {
        startIndex: insertIndex,
        endIndex: insertIndex + text.length
      },
      text_style: textStyle,
      fields: formFieldsString(textStyle)
    }
  }
  );
}


function insertPullQuote() {
  const ui = DocumentApp.getUi();
  const doc = DocumentApp.getActiveDocument();
  const documentId = doc.getId();

  const cursorPosition = detectCursorPosition(doc, documentId);
  if (cursorPosition.status == 'error') {
    ui.alert(cursorPosition.message);
    return 0;
  }

  const tableStartIndex = cursorPosition.endIndex;
  const numRows = 1;
  const numCols = 1;
  const requests = [];

  requests.push(
    {
      insertTable: {
        rows: numRows,
        columns: numCols,
        location: { index: tableStartIndex }
      }
    },
    {
      updateTableCellStyle: {
        tableRange: {
          tableCellLocation: {
            tableStartLocation: {
              index: tableStartIndex + 1
            },
          },
          rowSpan: numRows,
          columnSpan: numCols
        },

        tableCellStyle: {
          borderTop: tableStyle_TRANSPERENT_BORDER,
          borderBottom: tableStyle_TRANSPERENT_BORDER,
          borderLeft: tableStyle_TRANSPERENT_BORDER,
          borderRight: tableStyle_TRANSPERENT_BORDER,
          paddingTop: zeroMagnitude,
          paddingBottom: zeroMagnitude,
          paddingLeft: zeroMagnitude,
          paddingRight: zeroMagnitude
        },
        fields: 'borderTop,borderBottom,borderLeft,borderRight,paddingTop,paddingBottom,paddingLeft,paddingRight'
      }
    }
  );

  const text1 = '“Pull quote would go here, like this.”\n';
  addQuoteAPIrequest(requests, text1, tableStartIndex + 4, paragraphStyle_QUOTE_1, textStyle_QUOTE);
  const text2 = '– Author, 2022';
  addQuoteAPIrequest(requests, text2, tableStartIndex + 4 + text1.length, paragraphStyle_QUOTE_2, textStyle_QUOTE);

  requests.push(
    {
      deleteNamedRange: {
        name: cursorPosition.rangeName
      }
    },
    {
      deleteContentRange: {
        range: {
          startIndex: tableStartIndex - 1,
          endIndex: tableStartIndex,
        }
      }
    });

  Docs.Documents.batchUpdate({
    requests: requests
  }, documentId);
}

