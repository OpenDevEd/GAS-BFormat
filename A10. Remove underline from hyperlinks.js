// Old version 
//function removeUnderlineFromHyperlinks() {
//let ui = DocumentApp.getUi();
// try{
//   let requests = [];
//   let doc = DocumentApp.getActiveDocument();
//   let documentId = doc.getId();
//   let document = Docs.Documents.get(documentId);

//   document = Docs.Documents.get(documentId);
//   let bodyElements = document.body.content;

//   let allFootnotes = document.footnotes;
//   for (let footnoteId in allFootnotes) {
//     allFootnotes[footnoteId].content.forEach(function (content) {
//       content.paragraph.elements.forEach(function (item) {
//         helpRemoveUnderlineFromHyperlinks(requests, item, footnoteId);
//       });
//     });
//   }

//   for (let i in bodyElements) {
//     if (bodyElements[i].paragraph) {
//       bodyElements[i].paragraph.elements.forEach(function (item) {
//         if (item.textRun) {
//           helpRemoveUnderlineFromHyperlinks(requests, item);        
//         }

//       });
//     }
//   }

//   if (requests.length > 0) {
//     Docs.Documents.batchUpdate({
//       'requests': requests
//     }, documentId);
//   }
//   }
//   catch (error) {
//     ui.alert('Error in removeUnderlineFromHyperlinks: ' + error);
//     return 0;
//   }  
// }

// Old version removeUnderlineFromHyperlinks and defaultStyleReport use the function
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

// The function was copied from BZotero script
function removeUnderlineFromHyperlinks() {
  const ui = DocumentApp.getUi();
  try {
    const toDo = 'removeUnderlineFromHyperlinks';
    const doc = DocumentApp.getActiveDocument();
    changeAllBodyLinks(toDo, doc);
    changeAllFootnotesLinks(toDo, doc);
  }
  catch (error) {
    ui.alert('Error in removeUnderlineFromHyperlinks: ' + error);
  }
}

// The function was copied from BZotero script
function changeAllBodyLinks(toDo, doc) {
  const element = doc.getBody();
  changeAllLinks(element, toDo);
}

// The function was copied from BZotero script
function changeAllFootnotesLinks(toDo, doc) {
  const footnotes = doc.getFootnotes();
  let footnote, numChildren;
  for (let i in footnotes) {
    footnote = footnotes[i].getFootnoteContents();
    if (footnote == null) continue;
    numChildren = footnote.getNumChildren();
    for (let j = 0; j < numChildren; j++) {
      changeAllLinks(footnote.getChild(j), toDo);
    }
  }
}

// The function was copied from BZotero script (removeOpeninZoteroapp option was excluded)
function changeAllLinks(element, toDo) {

  let text, end, indices, partAttributes, numChildren, getIndexFlag;
  const elementType = String(element.getType());

  if (elementType == 'TEXT') {

    indices = element.getTextAttributeIndices();
    for (let i = 0; i < indices.length; i++) {
      partAttributes = element.getAttributes(indices[i]);
      if (partAttributes.LINK_URL) {

        getIndexFlag = false;

        if (toDo == 'removeUnderlineFromHyperlinks' && partAttributes.UNDERLINE) {
          getIndexFlag = true;
        }

        if (getIndexFlag === true) {
          if (i == indices.length - 1) {
            text = element.getText();
            end = text.length - 1;
          } else {
            end = indices[i + 1] - 1;
          }
          if (toDo == 'removeUnderlineFromHyperlinks') {
            element.setUnderline(indices[i], end, false);
          }
        }
      }
    }
  } else {
    const arrayTypes = ['BODY_SECTION', 'PARAGRAPH', 'LIST_ITEM', 'TABLE', 'TABLE_ROW', 'TABLE_CELL'];
    if (arrayTypes.includes(elementType)) {
      numChildren = element.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        changeAllLinks(element.getChild(i), toDo);
      }
    }
  }
}