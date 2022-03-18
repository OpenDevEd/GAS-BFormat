function formatBox() {
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
