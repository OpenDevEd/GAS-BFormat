const cmTOpt = 28.34645669;

// The function below was adapted from https://css-tricks.com/converting-color-spaces-in-javascript/#hex-to-rgb
// Convert Hex to RGB
function hexToRGB(h) {
  let r = 0, g = 0, b = 0;

  // 3 digits
  if (h.length == 4) {
    r = "0x" + h[1] + h[1];
    g = "0x" + h[2] + h[2];
    b = "0x" + h[3] + h[3];

  // 6 digits
  } else if (h.length == 7) {
    r = "0x" + h[1] + h[2];
    g = "0x" + h[3] + h[4];
    b = "0x" + h[5] + h[6];
  }

  return {red:+(r / 255), green:+(g / 255), blue:+(b / 255)};
}


// Find cursor position, create named range
// Return startIndex and endIndex of created named range
// insertPullQuote, insertTable, insertFigure1, insertFigure2 use the function
function detectCursorPosition(doc, documentId) {
  try {
    let rangeName;
    let cursor = doc.getCursor();
    if (cursor) {
      let rangeBuilder = doc.newRange();
      let insertedParagraph = cursor.insertText('~');

      rangeName = 'namedRange' + new Date().getTime();
      rangeBuilder.addElement(insertedParagraph);
      doc.addNamedRange(rangeName, rangeBuilder.build());
    } else {
      return {
        status: 'error',
        message: 'Cursor wasn\'t found.'
      };
    }

    doc.saveAndClose();

    // Find startIndex and endIndex of inserted paragraph (Docs API)
    let startEndIndex = getStartEndIndex(documentId, rangeName);
    if (startEndIndex.status == 'error') {
      return startEndIndex;
    } else {
      return {
        status: 'ok',
        startIndex: startEndIndex.startIndex,
        endIndex: startEndIndex.endIndex,
        rangeName: rangeName
      };
    }
  } catch (error) {
    return {
      status: 'error',
      message: 'Error in detectCursorPosition: ' + error
    };
  }
}

// Find selection or cursor position, create named range for elementType
// formatTextLikeH1, formatTable use the function
function getSelectionCreateNamedRange(doc, documentId, elementType) {
  try {
    let selection = doc.getSelection();
    let rangeBuilder = doc.newRange();
    let elements;
    if (selection) {
      elements = selection.getRangeElements();
    } else {
      let cursor = doc.getCursor();
      if (cursor) {
        elements = [cursor];
      } else {
        return {
          status: 'error',
          message: 'Please select ' + elementType.toLowerCase() + '.'
        };
      }
    }
    let found = false;
    let element;
    for (let i = 0; i < elements.length; i++) {
      element = elements[i].getElement();

      while (element.getType() != elementType && element.getType() != 'BODY_SECTION') {
        element = element.getParent();
      }
      if (element.getType() == elementType) {
        rangeBuilder.addElement(element);
        found = true;
      }
    }

    if (!found) {
      return {
        status: 'error',
        message: 'Please select ' + elementType.toLowerCase() + '.'
      };
    } else {

      let rangeName = 'namedRange' + new Date().getTime();
      doc.addNamedRange(rangeName, rangeBuilder.build());


      doc.saveAndClose();

      // Find startIndex and endIndex of selected named range (Docs API)
      let startEndIndex = getStartEndIndex(documentId, rangeName);
      if (startEndIndex.status == 'error') {
        return startEndIndex;
      } else {
        return {
          status: 'ok',
          startIndex: startEndIndex.startIndex,
          endIndex: startEndIndex.endIndex,
          rangeName: rangeName
        };
      }
    }
  } catch (error) {
    return {
      status: 'error',
      message: 'Error in getSelectionCreateNamedRange: ' + error
    };
  }
}

// Find startIndex and endIndex of selectedNamedRange
// detectCursorPosition, getSelectionCreateNamedRange use the function
function getStartEndIndex(documentId, selectedNamedRange) {
  try {
    let document = Docs.Documents.get(documentId);
    let startIndex;
    let endIndex;
    if (document.namedRanges) {
      if (document.namedRanges[selectedNamedRange]) {
        startIndex = document.namedRanges[selectedNamedRange].namedRanges[0].ranges[0].startIndex;
        endIndex = document.namedRanges[selectedNamedRange].namedRanges[0].ranges[0].endIndex;
      }
    }
    return {
      status: 'ok',
      startIndex: startIndex,
      endIndex: endIndex
    }
  } catch (error) {
    return {
      status: 'error',
      message: 'Error in getStartEndIndex: ' + error
    };
  }
}

// Get object that describe styling
// Return string "fields" for batchUpdate requests
// All functions that use Docs.Documents.batchUpdate use the function
function formFieldsString(object) {
  let string = '';
  let commaFlag = false;
  for (let key in object) {
    if (commaFlag === false) {
      commaFlag = true;
    } else {
      string += ',';
    }
    string += key;
  }
  if (string == '') string = '*';
  return string;
}

// Returns spaceBelow and spaceAbove from named styles of doc
// Not using now but it can be useful in future
function findNamedStyle(styles, namedStyleType) {
  const result = styles.filter(style => style.namedStyleType == namedStyleType);
  if (result.length == 0) {
    return {
      paragraphStyle: { spaceBelow: { magnitude: '' }, spaceAbove: { magnitude: '' } },
      textStyle: {}
    };
  } else {
    return result[0];
  }

}
