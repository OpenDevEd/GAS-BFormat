function formatBox() {
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

    tableStyle_ORANGE_BORDER.width.magnitude = 1.5;

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
            borderTop: tableStyle_ORANGE_BORDER,
            borderBottom: tableStyle_ORANGE_BORDER,
            borderLeft: tableStyle_ORANGE_BORDER,
            borderRight: tableStyle_ORANGE_BORDER,
          },
          fields: 'backgroundColor,borderBottom,borderLeft,borderRight,borderTop'
        }
      }
    );

    Docs.Documents.batchUpdate({
      requests: requests
    }, documentId);
  }
  catch (error) {
    ui.alert('Error in formatBox: ' + error);
    return 0;
  }
}
