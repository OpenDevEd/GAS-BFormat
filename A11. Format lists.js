function formatListsPart1(onlyLists = true, requests = [], body, document, documentId) {
  const ui = DocumentApp.getUi();
  try {
    if (onlyLists) {
      const doc = DocumentApp.getActiveDocument();
      body = doc.getBody();
      documentId = doc.getId();
      document = Docs.Documents.get(documentId);
    }

    let listItemsWarningText = '';
    let listItemsWarningTextBullets = '';
    let listItemsWarningTextColor = '';
    let glyphType;
    const expectedGlyphType = styles[ACTIVE_STYLE]['glyphType'];
    let lists = body.getListItems();
    lists.forEach(function (item) {
      const nestingLevel = item.getNestingLevel();
      const listId = item.getListId();
      item.setLineSpacing(1.15);

      if (document.lists[listId].listProperties.nestingLevels[nestingLevel].textStyle.foregroundColor) {
        if (document.lists[listId].listProperties.nestingLevels[nestingLevel].textStyle.foregroundColor.color.rgbColor) {
          const glyphColor = document.lists[listId].listProperties.nestingLevels[nestingLevel].textStyle.foregroundColor.color.rgbColor;
          if (!(glyphColor.blue == null && glyphColor.green == null && glyphColor.red == null)) {
            listItemsWarningTextColor += '\n' + item.getText();
          }
        }
      }

      glyphType = String(item.getGlyphType());
      if (nestingLevel == 0) {
        item.setIndentStart(36);
        /*if (['NUMBER', 'LATIN_UPPER', 'LATIN_LOWER', 'ROMAN_UPPER', 'ROMAN_LOWER', 'SQUARE_BULLET', 'HOLLOW_BULLET'].indexOf(glyphType) == -1) {
          item.setGlyphType(DocumentApp.GlyphType['SQUARE_BULLET']);
        }*/
        if (glyphType != expectedGlyphType){
          item.setGlyphType(DocumentApp.GlyphType[expectedGlyphType]);
        }
      } else if (nestingLevel == 1) {
        if (document.lists[listId].listProperties.nestingLevels[1].glyphSymbol != '–' && ['NUMBER', 'LATIN_UPPER', 'LATIN_LOWER', 'ROMAN_UPPER', 'ROMAN_LOWER'].indexOf(glyphType) == -1) {
          listItemsWarningTextBullets += '\n' + item.getText();
        }
        item.setIndentStart(72);
      }

    });

    if (listItemsWarningTextColor != '') {
      listItemsWarningText += '\nSet up black foreground colour for the bullet of the following list item(s):' + listItemsWarningTextColor;
    }
    if (listItemsWarningTextBullets != '') {
      listItemsWarningText += '\n\nSet up bullet – (en dash) for the following list item(s):' + listItemsWarningTextBullets;
    }

    if (listItemsWarningText != '') {
      ui.alert(listItemsWarningText);
    }

    if (onlyLists) {
      formatListsPart2(onlyLists = true, requests = [], document, documentId)
    }

  }
  catch (error) {
    ui.alert('Error in formatListsPart1: ' + error);
  }
}


function formatListsPart2(onlyLists = true, requests = [], document, documentId) {
  const ui = DocumentApp.getUi();
  try {
    let addRequest;
    const bodyElements = document.body.content;
    for (let i in bodyElements) {
      addRequest = false;
      if (bodyElements[i].paragraph) {
        // If paragraph has bullet
        if (bodyElements[i].paragraph.bullet) {

          // Check paragraph's style
          if (bodyElements[i].paragraph.paragraphStyle.spacingMode) {
            if (bodyElements[i].paragraph.paragraphStyle.spacingMode != 'NEVER_COLLAPSE') {
              bodyElements[i].paragraph.paragraphStyle.spacingMode = 'NEVER_COLLAPSE';
              addRequest = true;
            }
          } else {
            addRequest = true;
            bodyElements[i].paragraph.paragraphStyle.spacingMode = 'NEVER_COLLAPSE';
          }

          if (bodyElements[i].paragraph.paragraphStyle.spaceAbove) {
            if (!bodyElements[i].paragraph.paragraphStyle.spaceAbove.magnitude || bodyElements[i].paragraph.paragraphStyle.spaceAbove.magnitude != 0) {
              bodyElements[i].paragraph.paragraphStyle.spaceAbove.magnitude = 0;
              addRequest = true;
            }
          } else {
            addRequest = true;
            bodyElements[i].paragraph.paragraphStyle.spaceAbove = { magnitude: 0, unit: 'PT' };
          }
          if (bodyElements[i].paragraph.paragraphStyle.spaceBelow) {
            if (!bodyElements[i].paragraph.paragraphStyle.spaceBelow || bodyElements[i].paragraph.paragraphStyle.spaceBelow.magnitude != 10) {
              bodyElements[i].paragraph.paragraphStyle.spaceBelow.magnitude = 10;
              addRequest = true;
            }
          } else {
            addRequest = true;
            bodyElements[i].paragraph.paragraphStyle.spaceBelow = { magnitude: 10, unit: 'PT' };
          }
          if (addRequest) {
            requests.push({
              updateParagraphStyle: {
                paragraphStyle: bodyElements[i].paragraph.paragraphStyle,
                range: {
                  startIndex: bodyElements[i].startIndex,
                  endIndex: bodyElements[i].endIndex
                },
                fields: '*'
              }
            });
          }
          // End. Check paragraph's style

          // Check elements of paragraph
          bodyElements[i].paragraph.elements.forEach(function (item) {
            checkElementOfParagraph(requests, item, addRequest);
          });
          // End. Check elements of paragraph
        }
        // End. If paragraph has bullet
      }
    }

    if (requests.length > 0 && onlyLists) {
      Docs.Documents.batchUpdate({
        requests: requests
      }, documentId);
    }
  }
  catch (error) {
    ui.alert('Error in formatListsPart2: ' + error);
  }
}
