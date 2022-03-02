const paragraphStyle_LEFT_BORDER = {
  namedStyleType: 'NORMAL_TEXT',
  borderLeft: {
    width: {
      magnitude: 3,
      unit: 'PT'
    },
    padding: {
      magnitude: 6,
      unit: 'PT'
    },
    dashStyle: 'SOLID',
    color: {
      color: {
        rgbColor: hexToRGB(config_font_color)
      }
    }
  }
};


function leftBorderParagraph() {
  let ui = DocumentApp.getUi();
  let requests = [];
  try {
    let doc = DocumentApp.getActiveDocument();
    let documentId = doc.getId();

    // Create namedRange for selected paragraph
    let startEndIndex = getSelectionCreateNamedRange(doc, documentId, 'PARAGRAPH');
    if (startEndIndex.status == 'error') {
      ui.alert(startEndIndex.message);
      return 0;
    }

    let startIndex = startEndIndex.startIndex;
    let endIndex = startEndIndex.endIndex;
    let rangeName = startEndIndex.rangeName;


    let document = Docs.Documents.get(documentId);
    let bodyElements = document.body.content;


    for (let i in bodyElements) {
      // If body element contains paragraph
      if (bodyElements[i].paragraph) {
        if (bodyElements[i].paragraph.elements) {
          if (bodyElements[i].paragraph.elements[0]) {
            if (bodyElements[i].paragraph.elements[0].startIndex == startIndex) {
              requests.push({
                updateParagraphStyle: {
                  paragraphStyle: paragraphStyle_LEFT_BORDER,
                  range: {
                    startIndex: startIndex,
                    endIndex: endIndex
                  },
                  fields: formFieldsString(paragraphStyle_LEFT_BORDER)
                }
              });
              bodyElements[i].paragraph.elements.forEach(function (item) {
                if (item.textRun) {
                  if (item.textRun.textStyle) {
                    requests.push({
                      updateTextStyle: {
                        textStyle: item.textRun.textStyle,
                        range: {
                          startIndex: item.startIndex,
                          endIndex: item.endIndex
                        },
                        fields: '*'
                      }
                    });
                  }
                }
              });
            }
          }
        }
      }


      // }
    }
    // End. If body element contains paragraph

    requests.push({
      deleteNamedRange: {
        name: rangeName
      }
    });

    Docs.Documents.batchUpdate({
      requests: requests
    }, documentId);
  }
  catch (error) {
    ui.alert('Error in leftBorderParagraph: ' + error);
    return 0;
  }
}