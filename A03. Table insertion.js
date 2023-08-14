function insertTable2x2() {
  insertTable(2, 2);
}

function insertTable3x3() {
  insertTable(3, 3);
}

function insertTable4x4() {
  insertTable(4, 4);
}

const tableStyles = {
  textStyle_TOPIC_COLUMN_CELL: {
    fontSize: {
      magnitude: 12,
      unit: 'PT'
    },
    bold: true,
    weightedFontFamily: {
      fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
      weight: 400
    }
  },
  textStyle_ITEM_CELL: {
    fontSize: {
      magnitude: 12,
      unit: 'PT'
    },
    bold: false,
    weightedFontFamily: {
      fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
      weight: 400
    }
  }
};

const tableStyle_TRANSPERENT_BORDER = {
  width: {
    magnitude: 0,
    unit: 'PT'
  },
  dashStyle: 'SOLID',
  color: {
    color: {
      rgbColor: {}
    }
  }
};

const tableStyle_ORANGE_BORDER = {
  width: {
    magnitude: 1.0,
    unit: 'PT'
  },
  dashStyle: 'SOLID',
  color: {
    color: {
      rgbColor: hexToRGB(styles[ACTIVE_STYLE]['main_heading_font_color'])
    }
  },
};

const paragraphStyle_TABLE = {
  namedStyleType: 'NORMAL_TEXT',
  spaceAbove: { magnitude: styles[ACTIVE_STYLE]['paragraphSpacesInCell'], unit: 'PT' },
  spaceBelow: { magnitude: styles[ACTIVE_STYLE]['paragraphSpacesInCell'], unit: 'PT' },
  alignment: 'START',
};

const paragraphStyle_TABLE_HEADING = {
  namedStyleType: 'HEADING_6',
  spaceAbove: { magnitude: 10, unit: 'PT' },
  spaceBelow: { magnitude: 10, unit: 'PT' },
  alignment: 'START'
};

const textStyle_TABLE_HEADING_PART_1 = {
  foregroundColor: {
    color: {
      rgbColor: hexToRGB(styles[ACTIVE_STYLE]['customStyle']['h6']['FOREGROUND_COLOR'])
    }
  },
  fontSize: {
    magnitude: styles[ACTIVE_STYLE]['customStyle']['h6']['FONT_SIZE'],
    unit: 'PT'
  },
  bold: true,
  italic: false,
  weightedFontFamily: {
    fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
    weight: 400
  }
};

const textStyle_TABLE_HEADING_PART_2 = {
  foregroundColor: {
    color: {
      rgbColor: hexToRGB(styles[ACTIVE_STYLE]['customStyle']['h6']['FOREGROUND_COLOR'])
    }
  },
  fontSize: {
    magnitude: styles[ACTIVE_STYLE]['customStyle']['h6']['FONT_SIZE'],
    unit: 'PT'
  },
  bold: false,
  italic: true,
  weightedFontFamily: {
    fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
    weight: 400
  }
};

// Get number of rows and number of columns
// Return object that contains texts for cells, styles for cells, startIndex, endIndex for each cell
// insertTable use the function
function addTableParameters(numRows, numCols) {
  const table = [];
  let previousCellPos = 1;
  for (let row = 0; row < numRows; row++) {
    table.push([]);
    for (let column = 0; column < numCols; column++) {
      table[row].push({});
      // Describe cell's style
      if (row == 0 || column == 0) {
        if (row == 0 && column == 0) {
          table[row][column]['text'] = ' ';
        }
        table[row][column]['style'] = 'TOPIC_COLUMN_CELL';
        if (row != 0) {
          table[row][column]['text'] = 'Topic ' + row;
        }
        if (column != 0) {
          table[row][column]['text'] = 'Column ' + column;
        }

      } else {
        table[row][column]['style'] = 'ITEM_CELL';
        table[row][column]['text'] = 'Item';
      }
      if (column == 0) {
        table[row][column]['pos'] = previousCellPos + 3;
      } else {
        table[row][column]['pos'] = previousCellPos + 2;
      }
      previousCellPos = table[row][column]['pos'];
    }
  }
  return table;
}

// Get number of rows and number of columns
// Insert table using Doc API
// insertTable2x2, insertTable3x3, insertTable4x4 use the function
function insertTable(numRows, numCols) {
  const ui = DocumentApp.getUi();
  try {
    const doc = DocumentApp.getActiveDocument();
    const documentId = doc.getId();

    const cursorPosition = detectCursorPosition(doc, documentId);
    if (cursorPosition.status == 'error') {
      ui.alert(cursorPosition.message);
      return 0;
    }

    const tableStartIndex = cursorPosition.endIndex;

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
            borderBottom: tableStyle_ORANGE_BORDER,
            borderLeft: tableStyle_TRANSPERENT_BORDER,
            borderRight: tableStyle_TRANSPERENT_BORDER,
          },
          fields: 'borderBottom,borderLeft,borderRight'
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
            rowSpan: 1,
            columnSpan: numCols
          },

          tableCellStyle: {
            borderTop: tableStyle_TRANSPERENT_BORDER
          },
          fields: 'borderTop'
        }
      }
    );

    const table = addTableParameters(numRows, numCols);
    let textStyle;
    for (let row = numRows - 1; row >= 0; row--) {
      for (let col = numCols - 1; col >= 0; col--) {
        textStyle = tableStyles['textStyle_' + table[row][col]['style']];
        requests.push({
          insertText: {
            text: table[row][col]['text'],
            location: {
              index: tableStartIndex + table[row][col]['pos']
            },
          }
        },
          {
            updateParagraphStyle: {
              paragraphStyle: paragraphStyle_TABLE,
              range: {
                startIndex: tableStartIndex + table[row][col]['pos'],
                endIndex: tableStartIndex + table[row][col]['pos'] + table[row][col]['text'].length
              },
              fields: formFieldsString(paragraphStyle_TABLE)
            }
          },
          {
            updateTextStyle: {
              range: {
                startIndex: tableStartIndex + table[row][col]['pos'],
                endIndex: tableStartIndex + table[row][col]['pos'] + table[row][col]['text'].length
              },
              text_style: textStyle,
              fields: formFieldsString(textStyle)
            }
          }
        );
      }
    }

    requests.push({
      insertText: {
        text: 'Table X. Table title',
        location: {
          index: tableStartIndex
        }
      }
    },
      {
        updateParagraphStyle: {
          paragraphStyle: paragraphStyle_TABLE_HEADING,
          range: {
            startIndex: tableStartIndex,
            endIndex: tableStartIndex + 20
          },
          fields: formFieldsString(paragraphStyle_TABLE_HEADING)
        }
      },
      {
        updateTextStyle: {
          range: {
            startIndex: tableStartIndex,
            endIndex: tableStartIndex + 8
          },
          text_style: textStyle_TABLE_HEADING_PART_1,
          fields: formFieldsString(textStyle_TABLE_HEADING_PART_1)
        }
      },
      {
        updateTextStyle: {
          range: {
            startIndex: tableStartIndex + 8,
            endIndex: tableStartIndex + 20
          },
          text_style: textStyle_TABLE_HEADING_PART_2,
          fields: formFieldsString(textStyle_TABLE_HEADING_PART_2)
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
            startIndex: tableStartIndex - 1,
            endIndex: tableStartIndex,
          }

        }
      }
    );
 
    Docs.Documents.batchUpdate({
      requests: requests
    }, documentId);
  }
  catch (error) {
    ui.alert('Error in insertTable: ' + error);
    return 0;
  }
}