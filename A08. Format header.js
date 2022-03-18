const paragraphStyle_HEADING_SEC = {
  namedStyleType: "NORMAL_TEXT",
  borderBottom: {
    width: {
      magnitude: 1,
      unit: "PT"
    },
    color: {
      color: {
        rgbColor: hexToRGB(styles[ACTIVE_STYLE]['main_heading_font_color'])
      }
    },
    padding: {
      magnitude: 2,
      unit: 'PT'
    },
    dashStyle: 'SOLID'
  }
};

const textStyle_HEADING_SEC = {
  foregroundColor: {
    color: {
      rgbColor: {
        green: 0.0,
        red: 0.0,
        blue: 0.0
      }
    }
  },
  fontSize: {
    magnitude: styles[ACTIVE_STYLE]['header_font_size'],
    unit: 'PT'
  },
  bold: false,
  weightedFontFamily: {
    fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
    weight: 600
  }
};

function formatHeader(onlyHeader = true) {


  // Detect header text
  let headerText = '';
  if (styles[ACTIVE_STYLE]['title_position'] == 'header') {
    headerText = getDocumentTitle();
  }
  if (headerText == '') {
    headerText = styles[ACTIVE_STYLE]['header_text'];
  }
  // End. Detect header text


  
  const ui = DocumentApp.getUi();
  const requests = [];
  documentId = DocumentApp.getActiveDocument().getId();
  document = Docs.Documents.get(documentId);

  let result;
  if (document.documentStyle.defaultHeaderId == null) {
    result = insertHeader(requests, documentId, headerText);
  } else {
    result = updateHeader(requests, documentId, document);
  }

  if (result.status == 'error') {
    if (onlyHeader) {
      ui.alert(result.message);
      return 0;
    } else {
      return { status: 'error', message: result.message };
    }
  }

  return { status: 'ok', requests: requests };
}

function insertHeader(requests, documentId, headerText) {
  try {
    const requests2 = [];
    requests2.push(
      {
        createHeader: {
          type: 'DEFAULT'
        }
      }
    );
    Docs.Documents.batchUpdate({
      requests: requests2
    }, documentId);


    const document = Docs.Documents.get(documentId);

    let headerId;
    if (document.documentStyle.defaultHeaderId == null) {
      Logger.log('No Header');
    } else {
      headerId = document.documentStyle.defaultHeaderId;
    }


    const headerTextLength = headerText.length;
    requests.push(
      {
        updateDocumentStyle: {
          documentStyle: {
            useFirstPageHeaderFooter: true,
            pageNumberStart: 0,
            marginHeader: { magnitude: styles[ACTIVE_STYLE]['MARGIN_HEADER_cm'] * cmTOpt, unit: 'PT' }
          },
          fields: 'pageNumberStart,useFirstPageHeaderFooter,marginHeader'
        }
      },
      {
        insertText: {
          location: {
            segmentId: headerId,
            index: 0
          },
          text: headerText
        }
      },
      {
        updateParagraphStyle: {
          paragraphStyle: paragraphStyle_HEADING_SEC,
          range: {
            segmentId: headerId,
            startIndex: 0,
            endIndex: headerTextLength
          },
          fields: formFieldsString(paragraphStyle_HEADING_SEC)
        }
      },
      {
        updateTextStyle: {
          textStyle: textStyle_HEADING_SEC,
          range: {
            segmentId: headerId,
            startIndex: 0,
            endIndex: headerTextLength
          },
          fields: formFieldsString(textStyle_HEADING_SEC)
        }
      }
    );

    Docs.Documents.batchUpdate({ requests: requests }, documentId);
    return { status: 'ok' };
  }
  catch (error) {
    return { status: 'error', message: 'Error in insertHeader. ' + error };
  }
}

function updateHeader(requests, documentId, document) {
  try {
    const headerId = document.documentStyle.defaultHeaderId;

    // Check and remove firstPageHeader
    let firstPageHeaderId;
    if (document.documentStyle.firstPageHeaderId != null) {
      firstPageHeaderId = document.documentStyle.firstPageHeaderId;
      requests.push(
        {
          deleteHeader: {
            headerId: firstPageHeaderId
          }
        }
      );
    }
    // End. Check and remove firstPageHeader

    // Set up bottom border, different header for first page
    const endIndex = document.headers[headerId].content[0].endIndex;
    requests.push(
      {
        updateDocumentStyle: {
          documentStyle: {
            useFirstPageHeaderFooter: true,
            pageNumberStart: 0,
            marginHeader: { magnitude: styles[ACTIVE_STYLE]['MARGIN_HEADER_cm'] * cmTOpt, unit: 'PT' }
          },
          fields: 'pageNumberStart,useFirstPageHeaderFooter,marginHeader'
        }
      },
      {
        updateParagraphStyle: {
          paragraphStyle: paragraphStyle_HEADING_SEC,
          range: {
            segmentId: headerId,
            startIndex: 0,
            endIndex: endIndex
          },
          fields: formFieldsString(paragraphStyle_HEADING_SEC)
        }
      },
    );
    // End. Set up bottom border, different header for first page


    // Set up text style of header
    document.headers[headerId].content.forEach(function (item) {

      item.paragraph.elements.forEach(function (item) {

        item.textRun.textStyle['fontSize'] = { magnitude: styles[ACTIVE_STYLE]['header_font_size'], unit: 'PT' };
        item.textRun.textStyle['weightedFontFamily'] = { fontFamily: styles[ACTIVE_STYLE]['fontFamily'], weight: 600 };

        if (item.startIndex == null) {
          item.startIndex = 0;
        }

        requests.push({
          updateTextStyle: {
            textStyle: item.textRun.textStyle,
            range: {
              segmentId: headerId,
              startIndex: item.startIndex,
              endIndex: item.endIndex
            },
            fields: formFieldsString(item.textRun.textStyle)
          }
        });

      });
    });
    // End. Set up text style of header

    Docs.Documents.batchUpdate({ requests: requests }, documentId);
    return { status: 'ok' };
  }
  catch (error) {
    return { status: 'error', message: 'Error in updateHeader. ' + error };
  }
}