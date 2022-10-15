function insertBox() {
  const ui = DocumentApp.getUi();
  try {
    const textBoxH5 = 'Figure 3. Some text. Single cell box.';

    const lenTextBoxH5 = textBoxH5.length;
    
    const doc = DocumentApp.getActiveDocument();
    const documentId = doc.getId();

    const cursorPosition = detectCursorPosition(doc, documentId);
    if (cursorPosition.status == 'error') {
      ui.alert(cursorPosition.message);
      return 0;
    }

    const insertStartIndex = cursorPosition.endIndex;

    const requests = [];

tableStyle_ORANGE_BORDER.width.magnitude = 1.5;

    requests.push(
       {
        insertText: {
          text: textBoxH5,
          location: {
            index: insertStartIndex
          }
        }
      },     
      {
        insertTable: {
          rows: 1,
          columns: 1,
          location: { index: insertStartIndex + lenTextBoxH5}
        }
      },
      {
        updateTableCellStyle: {
          tableRange: {
            tableCellLocation: {
              tableStartLocation: {
                index: insertStartIndex + 1 + lenTextBoxH5
              },
            },
            rowSpan: 1,
            columnSpan: 1
          },

          tableCellStyle: {
            borderTop: tableStyle_ORANGE_BORDER,
            borderBottom: tableStyle_ORANGE_BORDER,
            borderLeft: tableStyle_ORANGE_BORDER,
            borderRight: tableStyle_ORANGE_BORDER,
          },
          fields: 'borderTop,borderBottom,borderLeft,borderRight'
        }
      },
      {
        updateParagraphStyle: {
          paragraphStyle: paragraphStyle_FIGURE_HEADING_5,
          range: {
            startIndex: insertStartIndex,
            endIndex: insertStartIndex + lenTextBoxH5
          },
          fields: formFieldsString(paragraphStyle_FIGURE_HEADING_5)
        }
      },
      {
        updateTextStyle: {
          range: {
            startIndex: insertStartIndex,
            endIndex: insertStartIndex + 9
          },
          text_style: textStyle_FIGURE_PART_1,
          fields: formFieldsString(textStyle_FIGURE_PART_1)
        }
      },
      {
        updateTextStyle: {
          range: {
            startIndex: insertStartIndex + 9,
            endIndex: insertStartIndex + lenTextBoxH5
          },
          text_style: textStyle_FIGURE_PART_2,
          fields: formFieldsString(textStyle_FIGURE_PART_2)
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
            startIndex: insertStartIndex - 1,
            endIndex: insertStartIndex,
          }

        }
      }
    );

    Docs.Documents.batchUpdate({
      requests: requests
    }, documentId);
  }
  catch (error) {
    ui.alert('Error in insertBox. ' + error);
  }
}