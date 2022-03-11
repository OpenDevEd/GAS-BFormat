let h1H2styles = {
  paragraphStyle_HEADING_1: {
    namedStyleType: 'HEADING_1'
  },
  textStyle_HEADING_1: {
    foregroundColor: {
      color: {
        rgbColor: hexToRGB(styles[getThisDocStyle()]['main_heading_font_color'])
      }
    },
    fontSize: {
      magnitude: 21,
      unit: 'PT'
    },
    bold: true,
    weightedFontFamily: {
      fontFamily: styles[getThisDocStyle()]['fontFamily'],
      weight: 400
    }
  },
  paragraphStyle_HEADING_2: {
    namedStyleType: 'HEADING_2',
    borderBottom: {
      width: {
        magnitude: 1,
        unit: 'PT'
      },
      padding: {
        magnitude: 2,
        unit: 'PT'
      },
      dashStyle: 'SOLID'
    }
  },
  textStyle_HEADING_2: {
    foregroundColor: {
      color: {
        rgbColor: {}
      }
    },
    fontSize: {
      magnitude: 16,
      unit: 'PT'
    },
    bold: true,
    weightedFontFamily: {
      fontFamily: styles[getThisDocStyle()]['fontFamily'],
      weight: 400
    }
  }
}

function formatTextLikeH1() {
  let ui = DocumentApp.getUi();
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

    h1H2styles.paragraphStyle_HEADING_1.namedStyleType = 'NORMAL_TEXT';
    let requests = [{
      updateParagraphStyle: {
        paragraphStyle: h1H2styles.paragraphStyle_HEADING_1,
        range: {
          startIndex: startIndex,
          endIndex: endIndex
        },
        fields: formFieldsString(h1H2styles.paragraphStyle_HEADING_1)
      }
    }, {
      updateTextStyle: {
        textStyle: h1H2styles.textStyle_HEADING_1,
        range: {
          startIndex: startIndex,
          endIndex: endIndex
        },
        fields: formFieldsString(h1H2styles.textStyle_HEADING_1)
      }
    },
    {
      deleteNamedRange: {
        name: rangeName
      }
    }];
    Docs.Documents.batchUpdate({
      requests: requests
    }, documentId);
  }
  catch (error) {
    ui.alert('Error in formatTextLikeH1: ' + error);
    return 0;
  }
}

// Set up style for all Heading_1 and Heading_2 paragraphs,
// normal text, footnotes, lists, remove underline from links
// Add/edit header and footer
function defaultStyleReport() {
  let ui = DocumentApp.getUi();
  try {

    const config_fontFamily = styles[getThisDocStyle()]['fontFamily'];

    let requests = [];
    let updateParagraphStyle;
    let doc = DocumentApp.getActiveDocument();
    let body = doc.getBody();
    let documentId = doc.getId();

    let document = Docs.Documents.get(documentId);

    // Set up body text (named style type NORMAL_TEXT) attributes
    let normalTextStyle = {};
    normalTextStyle[DocumentApp.Attribute.FONT_FAMILY] = config_fontFamily;
    normalTextStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
    normalTextStyle[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    normalTextStyle[DocumentApp.Attribute.SPACING_AFTER] = 10;
    normalTextStyle[DocumentApp.Attribute.LINE_SPACING] = 1.15;
    normalTextStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.NORMAL, normalTextStyle);
    // End. Set up body text (named style type NORMAL_TEXT) attributes

    // Set up heading 1 (named style type HEADING_1) attributes
    let heading1TextStyle = {};
    heading1TextStyle[DocumentApp.Attribute.FONT_FAMILY] = config_fontFamily;
    heading1TextStyle[DocumentApp.Attribute.FONT_SIZE] = 21;
    heading1TextStyle[DocumentApp.Attribute.SPACING_BEFORE] = 14;
    heading1TextStyle[DocumentApp.Attribute.SPACING_AFTER] = 10;
    heading1TextStyle[DocumentApp.Attribute.LINE_SPACING] = 1.15;
    heading1TextStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = styles[getThisDocStyle()]['main_heading_font_color'];
    heading1TextStyle[DocumentApp.Attribute.BOLD] = true;
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1, heading1TextStyle);
    // End. Set up heading 1 (named style type HEADING_1) attributes   

    // Set up heading 2 (named style type HEADING_2) attributes
    let heading2TextStyle = {};
    heading2TextStyle[DocumentApp.Attribute.FONT_FAMILY] = config_fontFamily;
    heading2TextStyle[DocumentApp.Attribute.FONT_SIZE] = 16;
    heading2TextStyle[DocumentApp.Attribute.SPACING_BEFORE] = 14;
    heading2TextStyle[DocumentApp.Attribute.SPACING_AFTER] = 8;
    heading2TextStyle[DocumentApp.Attribute.LINE_SPACING] = 1.15;
    heading2TextStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    heading2TextStyle[DocumentApp.Attribute.BOLD] = true;
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2, heading2TextStyle);
    // End. Set up heading 1 (named style type HEADING_1) attributes   

    // Set up heading 5(6) (named style type HEADING_5) attributes
    let heading5TextStyle = {};
    heading5TextStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING5, heading5TextStyle);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING6, heading5TextStyle);
    // End. Set up heading 5(6) (named style type HEADING_5) attributes

    // Set up footnotes attributes
    let footnoteStyle = {};
    footnoteStyle[DocumentApp.Attribute.FONT_FAMILY] = config_fontFamily;
    footnoteStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
    footnoteStyle[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    footnoteStyle[DocumentApp.Attribute.SPACING_AFTER] = 10;
    footnoteStyle[DocumentApp.Attribute.LINE_SPACING] = 1.15;
    let footnotes = body.getFootnotes();
    footnotes.forEach(function (item) {
      item.getFootnoteContents().getParagraphs().forEach(function (item) {
        item.setAttributes(footnoteStyle);
      });
    });
    // End. Set up footnotes attributes

    // Set up lists attributes part 1
    Logger.log('formatListsPart1');
    formatListsPart1(false, requests, body, document, documentId);

    // Set up 20pt after tables
    setSpaceAfterTables20pt(body);

    doc.saveAndClose();

    // Set up lists attributes part 2
    //formatListsPart2(false, requests, document, documentId);

    Logger.log('formatHeader');
    let resultHeader = formatHeader();
    if (resultHeader.status == 'error') {
      ui.alert(resultHeader.message);
    }

    Logger.log('formatFooter');
    let result = formatFooter();
    if (result.status == 'error') {
      ui.alert(result.message);
    }

    document = Docs.Documents.get(documentId);
    let bodyElements = document.body.content;

    let allFootnotes = document.footnotes;
    for (let footnoteId in allFootnotes) {
      allFootnotes[footnoteId].content.forEach(function (content) {
        content.paragraph.elements.forEach(function (item) {
          // Remove underline from hyperlinks
          helpRemoveUnderlineFromHyperlinks(requests, item, footnoteId);
        });
      });
    }


    let arrayH1H2 = [];
    let spaceAfterTableParagraph = true;
    let paragraphText = '';
    let emptyLineAfterTable;
    for (let i in bodyElements) {

      if (bodyElements[i].paragraph) {

        // Check paragraph after table
        emptyLineAfterTable = false;
        if (spaceAfterTableParagraph) {

          if (bodyElements[i].paragraph.elements) {
            bodyElements[i].paragraph.elements.forEach(function (item) {
              if (item.textRun) {
                if (item.textRun.content) {
                  paragraphText += item.textRun.content.trim();
                }
              }
            });
          }
          //Logger.log("AfterTableParagraph=" + paragraphText);
          if (paragraphText == '') {
            //Logger.log("Empty line=" + paragraphText);
            emptyLineAfterTable = true;
          }
          //  else {
          //   Logger.log("Not Empty line=" + paragraphText);
          // }

          spaceAfterTableParagraph = false;
          paragraphText = '';
        }
        // End. Check paragraph after table

        // If paragraph is list item, we set spacingMode = NEVER_COLLAPSE, spaceAbove = 10, spaceBelow = 0
        if (bodyElements[i].paragraph.bullet) {
          bodyElements[i].paragraph.paragraphStyle.spacingMode = 'NEVER_COLLAPSE';
        }
        // End. If paragraph is list item.

        let namedStyleType = bodyElements[i].paragraph.paragraphStyle.namedStyleType;
        let elements = bodyElements[i].paragraph.elements;
        let lastElement = elements.length - 1;

        if (namedStyleType == 'HEADING_1' || namedStyleType == 'HEADING_2') {
          // If paragraph has HEADING_1 or HEADING_2 named style, we push it in array arrayH1H2
          arrayH1H2.push({
            style: namedStyleType,
            startIndex: elements[0].startIndex,
            endIndex: elements[lastElement].endIndex
          });
        } else if (namedStyleType == 'NORMAL_TEXT') {
          // If paragraph has NORMAL_TEXT named style

          // Check paragraph's style
          let spaceBelow = bodyElements[i].paragraph.paragraphStyle.spaceBelow;
          let spaceAbove = bodyElements[i].paragraph.paragraphStyle.spaceAbove;

          let itIsExtractedQuote = false;

          if (spaceBelow && spaceAbove) {
            if (spaceBelow.magnitude && spaceAbove.magnitude) {
              if (spaceBelow.magnitude == 20 && spaceAbove.magnitude == 20) {
                itIsExtractedQuote = true;
              }
            }
          }

          if (spaceBelow) {
            if (spaceBelow.magnitude) {
              if (spaceBelow.magnitude != 10) {
                spaceBelow.magnitude = 10;
                updateParagraphStyle = true;
              }
            } else {
              spaceBelow.magnitude = 10;
              updateParagraphStyle = true;
            }
          }


          if (spaceAbove) {
            if (spaceAbove.magnitude != 0) {
              spaceAbove.magnitude = 0;
              updateParagraphStyle = true;
            }
          }



          // Logger.log('elements[0].startIndex' + elements[0].startIndex);
          // Logger.log('elements[lastElement].endIndex' + elements[lastElement].endIndex)
          if (updateParagraphStyle && !itIsExtractedQuote && !emptyLineAfterTable) {
            requests.push({
              updateParagraphStyle: {
                paragraphStyle: bodyElements[i].paragraph.paragraphStyle,
                range: {
                  startIndex: elements[0].startIndex,
                  endIndex: elements[lastElement].endIndex
                },
                fields: '*'
              }
            });
          }
          // End. Check paragraph's style

          // Check elements of paragraph
          if (!itIsExtractedQuote && !emptyLineAfterTable) {
            bodyElements[i].paragraph.elements.forEach(function (item) {
              checkElementOfParagraph(requests, item, updateParagraphStyle);
            });
          }
          // End. Check elements of paragraph
          // End. If paragraph has NORMAL_TEXT named style
        }
      } else if (bodyElements[i].table) {
        spaceAfterTableParagraph = true;
      }
    }

    // Use data from arrayH1H2 to add requests for Docs.Documents.batchUpdate
    for (let i in arrayH1H2) {
      requests.push({
        updateParagraphStyle: {
          paragraphStyle: h1H2styles['paragraphStyle_' + arrayH1H2[i].style],
          range: {
            startIndex: arrayH1H2[i].startIndex,
            endIndex: arrayH1H2[i].endIndex
          },
          fields: formFieldsString(h1H2styles['paragraphStyle_' + arrayH1H2[i].style])
        }
      }, {
        updateTextStyle: {
          textStyle: h1H2styles['textStyle_' + arrayH1H2[i].style],
          range: {
            startIndex: arrayH1H2[i].startIndex,
            endIndex: arrayH1H2[i].endIndex
          },
          fields: formFieldsString(h1H2styles['textStyle_' + arrayH1H2[i].style])
        }
      });
    }
    // End. Use data from arrayH1H2 to add requests for Docs.Documents.batchUpdate

    // Set up lists attributes part 2
    //  formatListsPart2(false, requests, document, documentId);


    let l = requests.length;
    for (let i = l - 1; i >= 0; i--) {
      if (requests[i].updateParagraphStyle) {
        if (!requests[i].updateParagraphStyle.paragraphStyle) {
          Logger.log(i + 'splice' + JSON.stringify(requests[i]));
          requests.splice(i, 1);
        }
      }
    }

    if (requests.length > 0) {
      Docs.Documents.batchUpdate({
        requests: requests
      }, documentId);
    } else {
      //  ui.alert('Nothing to change!');
    }
  } catch (error) {
    Logger.log('Error in defaultStyleReport: ' + error);
    ui.alert('Error in defaultStyleReport: ' + error);
  }
}

// Check fontSize and fontFamily of element of paragraph
// defaultStyleReport and formatListsPart2 use the function
function checkElementOfParagraph(requests, item, updateParagraphStyle) {

  const config_fontFamily = styles[getThisDocStyle()]['fontFamily'];

  let normalTextForegroundColor, likeH1textColor;
  likeH1textColor = false;
  wrongFontSize = false;
  if (item.textRun) {

    // Remove underline from hyperlinks
    helpRemoveUnderlineFromHyperlinks(requests, item);

    // Check fontSize
    if (item.textRun.textStyle.fontSize) {
      if (item.textRun.textStyle.fontSize.magnitude != 12) {

        // Check whether the text has orange color that was set up by function formatTextLikeH1
        normalTextForegroundColor = item.textRun.textStyle.foregroundColor;
        if (normalTextForegroundColor) {
          if (normalTextForegroundColor.color) {
            if (normalTextForegroundColor.color.rgbColor) {
              if (normalTextForegroundColor.color.rgbColor.green == h1H2styles.textStyle_HEADING_1.foregroundColor.color.rgbColor.green
                && normalTextForegroundColor.color.rgbColor.red == h1H2styles.textStyle_HEADING_1.foregroundColor.color.rgbColor.red) {
                likeH1textColor = true;
              }
            }
          }
        }
        // End. Check whether the text has orange color that was set up by function formatTextLikeH1

        // Note! We don't change fontSize if orange color and fontSize 21 PT were set up by function formatTextLikeH1
        if (item.textRun.textStyle.fontSize.magnitude != h1H2styles.textStyle_HEADING_1.fontSize.magnitude && likeH1textColor === false) {
          wrongFontSize = true;
          item.textRun.textStyle.fontSize.magnitude = 12;
        }

      }
    }
    // End. Check fontSize

    // Check fontFamily
    if (item.textRun.textStyle.weightedFontFamily) {
      if (item.textRun.textStyle.weightedFontFamily.fontFamily != config_fontFamily) {
        wrongFontSize = true;
        item.textRun.textStyle.weightedFontFamily.fontFamily = config_fontFamily;
      }
    }
    // End. Check fontFamily

    // Add updateTextStyle request
    if (wrongFontSize === true || updateParagraphStyle === true) {
      if (item.startIndex == null) {
        item.startIndex = 0;
      }
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
    // End. Add updateTextStyle request
  }
}
