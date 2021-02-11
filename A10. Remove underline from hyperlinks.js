function removeUnderlineFromHyperlinks() {
  let ui = DocumentApp.getUi();
  try{
  let requests = [];
  let doc = DocumentApp.getActiveDocument();
  let documentId = doc.getId();
  let document = Docs.Documents.get(documentId);

  document = Docs.Documents.get(documentId);
  let bodyElements = document.body.content;

  let allFootnotes = document.footnotes;
  for (let footnoteId in allFootnotes) {
    allFootnotes[footnoteId].content.forEach(function (content) {
      content.paragraph.elements.forEach(function (item) {
        helpRemoveUnderlineFromHyperlinks(requests, item, footnoteId);
      });
    });
  }

  for (let i in bodyElements) {
    if (bodyElements[i].paragraph) {
      bodyElements[i].paragraph.elements.forEach(function (item) {
        if (item.textRun) {
          helpRemoveUnderlineFromHyperlinks(requests, item);        
        }

      });
    }
  }

  if (requests.length > 0) {
    Docs.Documents.batchUpdate({
      'requests': requests
    }, documentId);
  }
  }
  catch (error) {
    ui.alert('Error in removeUnderlineFromHyperlinks: ' + error);
    return 0;
  }  
}

function helpRemoveUnderlineFromHyperlinks(requests, item, segmentId) {
  if (item.textRun.textStyle) {
    if (item.textRun.textStyle.link) {
      item.textRun.textStyle.underline = false;
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
      if (segmentId) {
        requests[requests.length - 1].updateTextStyle.range.segmentId = segmentId;
      }
    }
  }
}
