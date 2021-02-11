const zeroMagnitude = {
  magnitude: 0,
  unit: 'PT'
};

let paragraphStyle_QUOTE_1 = {
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
        rgbColor: {
          green: 0.36078432,
          red: 1.0
        }
      }
    },
  }
};

let paragraphStyle_QUOTE_2 = {
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
        rgbColor: {
          green: 0.36078432,
          red: 1.0
        }
      }
    },
  }
};

let textStyle_QUOTE = {
  foregroundColor: {
    color: {
      rgbColor: {
        green: 0.36078432,
        red: 1.0
      }
    }
  },
  fontSize: {
    magnitude: 14,
    unit: 'PT'
  },
  bold: true,
  weightedFontFamily: {
    fontFamily: 'Montserrat',
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
  let ui = DocumentApp.getUi();
  let doc = DocumentApp.getActiveDocument();
  let documentId = doc.getId();

  let cursorPosition = detectCursorPosition(doc, documentId);
  if (cursorPosition.status == 'error') {
    ui.alert(cursorPosition.message);
    return 0;
  }

  let tableStartIndex = cursorPosition.endIndex;
  let numRows = 1;
  let numCols = 1;
  let requests = [];

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

  let text1 = '“Pull quote would go here, like this.”\n';
  addQuoteAPIrequest(requests, text1, tableStartIndex + 4, paragraphStyle_QUOTE_1, textStyle_QUOTE);
  let text2 = '– Author, 2021';
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

