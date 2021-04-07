function formatTableBasic() {

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

    let requests = [];

    const numRows = fTable.table.rows;
    const numCols = fTable.table.columns;

    const cellPadding = {
      magnitude: 4.25,
      unit: 'PT'
    }

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
            contentAlignment: 'MIDDLE',
            paddingLeft: cellPadding,
            paddingRight: cellPadding,
            paddingTop: cellPadding,
            paddingBottom: cellPadding
          },
          fields: 'contentAlignment,paddingLeft,paddingRight,paddingTop,paddingBottom'
        }
      },      
      {
        deleteNamedRange: {
          name: namedRange.rangeName
        }
      }
    );

    Docs.Documents.batchUpdate({
      requests: requests
    }, documentId);
  }
  catch (error) {
    ui.alert('Error in formatTableBasic: ' + error);
    return 0;
  }
}