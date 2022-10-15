// Formats table using default Report styles
// but doesn't update styles of paragraphs and texts
function formatTableNoBold() {
  formatTable(updateParagraphTextStyle = false);
}

// Format table using default Report styles
// See styles in A03. Table insertion.gs
function formatTable(updateParagraphTextStyle = true) {
  const ui = DocumentApp.getUi();
  try {
    const doc = DocumentApp.getActiveDocument();
    const documentId = doc.getId();

    // Create namedRange for selected table
    const namedRange = getSelectionCreateNamedRange(doc, documentId, 'TABLE');
    if (namedRange.status == 'error') {
      ui.alert(namedRange.message);
      return 0;
    }

    const tableStartIndex = namedRange.startIndex;
    const tableEndIndex = namedRange.endIndex;


    const document = Docs.Documents.get(documentId);

    const bodyElements = document.body.content;

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

    const requests = [];

    const numRows = fTable.table.rows;
    const numCols = fTable.table.columns;

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
            borderBottom: tableStyle_ORANGE_BORDER,
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

    // Update paragraph style and text style
    if (updateParagraphTextStyle) {
      let textStyle;
      for (let row = 0; row < numRows; row++) {
        for (let col = 0; col < numCols; col++) {
          if (row == 0 || col == 0) {
            textStyle = tableStyles.textStyle_TOPIC_COLUMN_CELL;
          } else {
            textStyle = tableStyles.textStyle_ITEM_CELL;
          }


          const cellStartIndex = fTable.table.tableRows[row].tableCells[col].startIndex;
          const cellEndIndex = fTable.table.tableRows[row].tableCells[col].endIndex;

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
    }
    // End. Update paragraph style and text style

    Docs.Documents.batchUpdate({
      requests: requests
    }, documentId);
  }
  catch (error) {
    ui.alert('Error in formatTable: ' + error);
    return 0;
  }
}
