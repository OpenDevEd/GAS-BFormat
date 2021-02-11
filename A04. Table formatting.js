// Format table using default Report styles
// See styles in A03. Table insertion.gs
function formatTable() {
  let ui = DocumentApp.getUi();
  try {
    let doc = DocumentApp.getActiveDocument();
    let documentId = doc.getId();

    // Create namedRange for selected table
    let namedRange = getSelectionCreateNamedRange(doc, documentId, 'TABLE');
    if (namedRange.status == 'error') {
      ui.alert(namedRange.message);
      return 0;
    }

    let tableStartIndex = namedRange.startIndex;
    let tableEndIndex = namedRange.endIndex;


    let document = Docs.Documents.get(documentId);

    let bodyElements = document.body.content;

    let fTable;
    for (let i in bodyElements) {
      if (bodyElements[i].table) {
        if (bodyElements[i].startIndex == tableStartIndex) {
          if (bodyElements[i].endIndex == tableEndIndex) {
            fTable = bodyElements[i];
            break;
          }
        }
      }
    }

    let requests = [];

    let numRows = fTable.table.rows;
    let numCols = fTable.table.columns;

    requests.push(
      {
        updateTableCellStyle: {
          tableRange: {
            tableCellLocation: {
              tableStartLocation: {
                index: tableStartIndex
              },
            },
            rowSpan: numRows,
            columnSpan: numCols
          },

          tableCellStyle: {
            borderBottom: {
              width: {
                magnitude: 1.0,
                unit: 'PT'
              },
              dashStyle: 'SOLID',
              color: {
                color: {
                  rgbColor: {
                    'green': 0.36078432,
                    red: 1.0
                  }
                }
              },
            },
            borderLeft: tableStyle_TRANSPERENT_BORDER,
            borderRight: tableStyle_TRANSPERENT_BORDER,
          },
          fields: 'backgroundColor,borderBottom,borderLeft,borderRight'
        }
      },
      {
        updateTableCellStyle: {
          tableRange: {
            tableCellLocation: {
              tableStartLocation: {
                index: tableStartIndex
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

    let textStyle;
    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numCols; col++) {
        if (row == 0 || col == 0) {
          textStyle = tableStyles.textStyle_TOPIC_COLUMN_CELL;
        } else {
          textStyle = tableStyles.textStyle_ITEM_CELL;
        }


        let cellStartIndex = fTable.table.tableRows[row].tableCells[col].startIndex;
        let cellEndIndex = fTable.table.tableRows[row].tableCells[col].endIndex;

        requests.push(
          {
            updateParagraphStyle: {
              paragraphStyle: paragraphStyle_TABLE,
              range: {
                startIndex: cellStartIndex,
                endIndex: cellEndIndex
              },
              fields: formFieldsString(paragraphStyle_TABLE)
            }
          },
          {
            updateTextStyle: {
              range: {
                startIndex: cellStartIndex,
                endIndex: cellEndIndex
              },
              text_style: textStyle,
              fields: formFieldsString(textStyle)
            }
          }
        );
      }
    }

    Docs.Documents.batchUpdate({
      requests: requests
    }, documentId);
  }
  catch (error) {
    ui.alert('Error in formatTable: ' + error);
    return 0;
  }
}




