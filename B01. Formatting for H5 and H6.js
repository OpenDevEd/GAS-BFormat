function reformatHeadings5and6() {
  try {
    const requests = [];
    const doc = DocumentApp.getActiveDocument();
    const documentId = doc.getId();
    const document = Docs.Documents.get(documentId);

    //  document = Docs.Documents.get(documentId);
    const bodyElements = document.body.content;

    for (let i in bodyElements) {
      // If body element contains table
      if (bodyElements[i].table) {
        if (bodyElements[i].table.tableRows) {
          for (let j in bodyElements[i].table.tableRows) {
            if (bodyElements[i].table.tableRows[j].tableCells) {
              for (let k in bodyElements[i].table.tableRows[j].tableCells) {
                if (bodyElements[i].table.tableRows[j].tableCells[k].content) {
                  for (let l in bodyElements[i].table.tableRows[j].tableCells[k].content) {
                    if (bodyElements[i].table.tableRows[j].tableCells[k].content[l].paragraph) {
                      detectHeadings5and6(requests, bodyElements[i].table.tableRows[j].tableCells[k].content[l].paragraph);
                    }
                  }
                }
              }
            }
          }

        }
      }
      // End. If body element contains table

      // If body element contains paragraph
      if (bodyElements[i].paragraph) {
        detectHeadings5and6(requests, bodyElements[i].paragraph);
      }
      // End. If body element contains paragraph
    }

    if (requests.length > 0) {
      Docs.Documents.batchUpdate({
        requests: requests
      }, documentId);
    }
  }
  catch (error) {
    ui.alert('Error in reformatHeadings5and6: ' + error);
    return 0;
  }
}

function detectHeadings5and6(requests, paragraph) {
  let paragraphText = '';
  let start, end;
  const namedStyleType = paragraph.paragraphStyle.namedStyleType;

  if ((namedStyleType == 'HEADING_5' || namedStyleType == 'HEADING_6') && paragraph.elements) {
    paragraph.elements.forEach(function (item) {
      if (item.textRun) {
        if (item.textRun.content) {
          paragraphText += item.textRun.content;
          if (start == null) {
            start = item.startIndex;
          }
          end = item.endIndex;
        }
      }
    });

    //const checkTable = /^Table (\d+|X)\./.exec(paragraphText);
    const checkTable = /^Table (\d+|X)\.?\d*\. /.exec(paragraphText);
    if (checkTable != null) {
      updateFigureXBoxXTableXStyle(requests, start, checkTable[0].length, end, paragraphStyle_TABLE_HEADING, textStyle_TABLE_HEADING_PART_1, textStyle_TABLE_HEADING_PART_2);
    } else {
      //const checkBoxFigure = /^(Figure|Box) (\d+|X)\./.exec(paragraphText);
      const checkBoxFigure = /^(Figure|Box) (\d+|X)\.?\d*\. /.exec(paragraphText);
      if (checkBoxFigure != null) {
        updateFigureXBoxXTableXStyle(requests, start, checkBoxFigure[0].length, end, paragraphStyle_FIGURE_HEADING_5, textStyle_FIGURE_PART_1, textStyle_FIGURE_PART_2);
      }
    }
  }
}

function updateFigureXBoxXTableXStyle(requests, startIndex, tableXLength, endIndex, paragraphStyle, textStyle1, textStyle2) {
  requests.push(
    {
      updateParagraphStyle: {
        paragraphStyle: paragraphStyle,
        range: {
          startIndex: startIndex,
          endIndex: endIndex
        },
        fields: formFieldsString(paragraphStyle)
      }
    },
    {
      updateTextStyle: {
        range: {
          startIndex: startIndex,
          endIndex: startIndex + tableXLength
        },
        text_style: textStyle1,
        fields: formFieldsString(textStyle1)
      }
    },
    {
      updateTextStyle: {
        range: {
          startIndex: startIndex + tableXLength,
          endIndex: endIndex
        },
        text_style: textStyle2,
        fields: formFieldsString(textStyle2)
      }
    }
  );
}
