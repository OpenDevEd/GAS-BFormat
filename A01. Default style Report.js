const h1H2styles = {
  paragraphStyle_HEADING_1: {
    namedStyleType: 'HEADING_1'
  },
  textStyle_HEADING_1: {
    foregroundColor: {
      color: {
        rgbColor: hexToRGB(styles[ACTIVE_STYLE]['main_heading_font_color'])
      }
    },
    fontSize: {
      magnitude: styles[ACTIVE_STYLE]['customStyle']['h1']['FONT_SIZE'],
      unit: 'PT'
    },
    bold: true,
    weightedFontFamily: {
      fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
      weight: 400
    }
  },

  paragraphStyle_HEADING_2: {
    namedStyleType: 'HEADING_2',
    /* borderBottom: {
       width: {
         magnitude: 1,
         unit: 'PT'
       },
       padding: {
         magnitude: 2,
         unit: 'PT'
       },
       dashStyle: 'SOLID'
     } */
  },
  textStyle_HEADING_2: {
    foregroundColor: {
      color: {
        rgbColor: {}
      }
    },
    fontSize: {
      magnitude: styles[ACTIVE_STYLE]['customStyle']['h2']['FONT_SIZE'],
      unit: 'PT'
    },
    bold: true,
    weightedFontFamily: {
      fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
      weight: 400
    }
  },

  paragraphStyle_HEADING_3: {
    namedStyleType: 'HEADING_3'
  },
  textStyle_HEADING_3: {},

  paragraphStyle_HEADING_4: {
    namedStyleType: 'HEADING_4'
  },
  textStyle_HEADING_4: {},

  paragraphStyle_HEADING_5: {
    namedStyleType: 'HEADING_5'
  },
  textStyle_HEADING_5: {},

  paragraphStyle_HEADING_6: {
    namedStyleType: 'HEADING_6'
  },
  textStyle_HEADING_6: {},
}

function addTextStyleHeading() {
  for (let i = 1; i <= 6; i++) {
    h1H2styles['textStyle_HEADING_' + i] = {
      /*      foregroundColor: {
              color: {
                rgbColor: hexToRGB(styles[ACTIVE_STYLE]['customStyle']['h'+i]['FOREGROUND_COLOR'])
              }
            },
            fontSize: {
              magnitude: styles[ACTIVE_STYLE]['customStyle']['h'+i]['FONT_SIZE'],
              unit: 'PT'
            }, */
      weightedFontFamily: {
        fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
        weight: 400
      }
    };

    if (styles[ACTIVE_STYLE]['customStyle'].hasOwnProperty('h' + i)) {
      if (styles[ACTIVE_STYLE]['customStyle']['h' + i].hasOwnProperty('ITALIC')) {
        h1H2styles['textStyle_HEADING_' + i].italic = styles[ACTIVE_STYLE]['customStyle']['h' + i]['ITALIC'];
      }
      if (styles[ACTIVE_STYLE]['customStyle']['h' + i].hasOwnProperty('BOLD') && styles[ACTIVE_STYLE]['customStyle']['h' + i]['BOLD'] === false) {
        h1H2styles['textStyle_HEADING_' + i].bold = styles[ACTIVE_STYLE]['customStyle']['h' + i]['BOLD'];
      } else {
        h1H2styles['textStyle_HEADING_' + i].bold = true;
      }
      if (styles[ACTIVE_STYLE]['customStyle']['h' + i].hasOwnProperty('FOREGROUND_COLOR')) {
        h1H2styles['textStyle_HEADING_' + i].foregroundColor = {
          color: {
            rgbColor: hexToRGB(styles[ACTIVE_STYLE]['customStyle']['h' + i]['FOREGROUND_COLOR'])
          }
        };
      }
      if (styles[ACTIVE_STYLE]['customStyle']['h' + i].hasOwnProperty('FONT_SIZE')) {
        h1H2styles['textStyle_HEADING_' + i].fontSize = {
          magnitude: styles[ACTIVE_STYLE]['customStyle']['h' + i]['FONT_SIZE'],
          unit: 'PT'
        }
      }
    }
  }
  //Logger.log(h1H2styles);
}


function formatTextLikeH1() {
  const ui = DocumentApp.getUi();
  try {
    const doc = DocumentApp.getActiveDocument();
    const documentId = doc.getId();

    // Create namedRange for selected paragraph
    const startEndIndex = getSelectionCreateNamedRange(doc, documentId, 'PARAGRAPH');
    if (startEndIndex.status == 'error') {
      ui.alert(startEndIndex.message);
      return 0;
    }

    const startIndex = startEndIndex.startIndex;
    const endIndex = startEndIndex.endIndex;
    const rangeName = startEndIndex.rangeName;

    h1H2styles.paragraphStyle_HEADING_1.namedStyleType = 'NORMAL_TEXT';
    const requests = [{
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

function createStyle(defaultStyle, customStyle) {
  for (let heading in customStyle) {
    if (defaultStyle.hasOwnProperty(heading) === false) {
      defaultStyle[heading] = {};
    }
    for (let attribute in customStyle[heading]) {
      if (attribute != 'FOREGROUND_COLOR') {
        defaultStyle[heading][attribute] = customStyle[heading][attribute];
      }
    }
  }
  for (let heading in defaultStyle) {
    if (defaultStyle[heading].hasOwnProperty() === false) {
      defaultStyle[heading]['FONT_FAMILY'] = styles[ACTIVE_STYLE]['fontFamily'];
    }
  }
  return defaultStyle;
}

// Set up style for all Heading_1 and Heading_2 paragraphs,
// normal text, footnotes, lists, remove underline from links
// Add/edit header and footer
function defaultStyleReport() {
  const ui = DocumentApp.getUi();
  try {

    const { status } = findAndMarkBrokenLinksOptimised();
    if (status === 'error') {
      return { status };
    }

    const defaultStyle = {//, FOREGROUND_COLOR: '#000000'
      normalText: { FONT_SIZE: 12, SPACING_BEFORE: 0, SPACING_AFTER: 10, LINE_SPACING: 1.15 },
      // , FOREGROUND_COLOR: '#FF5C00'
      h1: { FONT_SIZE: 21, SPACING_BEFORE: 14, SPACING_AFTER: 10, LINE_SPACING: 1.15 },
      //, FOREGROUND_COLOR: '#000000' 
      h2: { FONT_SIZE: 16, SPACING_BEFORE: 14, SPACING_AFTER: 8, LINE_SPACING: 1.15 },
      h3: {},
      h4: {},
      h5: {},
      h6: {},
      // h3: { FOREGROUND_COLOR: '#000000' },
      // h4: { FOREGROUND_COLOR: '#000000' },
      // h5: { FOREGROUND_COLOR: '#000000' },
      // h6: { FOREGROUND_COLOR: '#000000' },
      footnote: { FONT_SIZE: 10, SPACING_BEFORE: 0, SPACING_AFTER: 0, LINE_SPACING: 1 }
    };
    const customStyle = createStyle(defaultStyle, styles[ACTIVE_STYLE]['customStyle']);

    const config_fontFamily = styles[ACTIVE_STYLE]['fontFamily'];

    const requests = [];
    let updateParagraphStyle;
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const documentId = doc.getId();

    let document = Docs.Documents.get(documentId);

    body.setPageWidth(styles[ACTIVE_STYLE]['pageWidth_cm'] * cmTOpt);
    body.setPageHeight(styles[ACTIVE_STYLE]['pageHeight_cm'] * cmTOpt);

    setMarginsHelper(body);

    // Set up body text (named style type NORMAL_TEXT) attributes
    /*const normalTextStyle = {};
    normalTextStyle[DocumentApp.Attribute.FONT_FAMILY] = config_fontFamily;
    normalTextStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
    normalTextStyle[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    normalTextStyle[DocumentApp.Attribute.SPACING_AFTER] = 10;
    normalTextStyle[DocumentApp.Attribute.LINE_SPACING] = 1.15;
    normalTextStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.NORMAL, normalTextStyle);*/
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.NORMAL, customStyle['normalText']);
    // End. Set up body text (named style type NORMAL_TEXT) attributes

    // Set up heading 1 (named style type HEADING_1) attributes
    /* const heading1TextStyle = {};
    heading1TextStyle[DocumentApp.Attribute.FONT_FAMILY] = config_fontFamily;
    heading1TextStyle[DocumentApp.Attribute.FONT_SIZE] = 21;
    heading1TextStyle[DocumentApp.Attribute.SPACING_BEFORE] = 14;
    heading1TextStyle[DocumentApp.Attribute.SPACING_AFTER] = 10;
    heading1TextStyle[DocumentApp.Attribute.LINE_SPACING] = 1.15;
    heading1TextStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = styles[ACTIVE_STYLE]['main_heading_font_color'];
    heading1TextStyle[DocumentApp.Attribute.BOLD] = true; */
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING1, customStyle['h1']);
    // End. Set up heading 1 (named style type HEADING_1) attributes   

    // Set up heading 2 (named style type HEADING_2) attributes
    /*const heading2TextStyle = {};
    heading2TextStyle[DocumentApp.Attribute.FONT_FAMILY] = config_fontFamily;
    heading2TextStyle[DocumentApp.Attribute.FONT_SIZE] = 16;
    heading2TextStyle[DocumentApp.Attribute.SPACING_BEFORE] = 14;
    heading2TextStyle[DocumentApp.Attribute.SPACING_AFTER] = 8;
    heading2TextStyle[DocumentApp.Attribute.LINE_SPACING] = 1.15;
    heading2TextStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    heading2TextStyle[DocumentApp.Attribute.BOLD] = true;*/
    //body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2, heading2TextStyle);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING2, customStyle['h2']);
    // End. Set up heading 1 (named style type HEADING_1) attributes  

    if (customStyle.hasOwnProperty('h3')) {
      body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING3, customStyle['h3']);
    }

    if (customStyle.hasOwnProperty('h4')) {
      body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING4, customStyle['h4']);
    }

    // Set up heading 5(6) (named style type HEADING_5) attributes
    /*const heading5TextStyle = {};
    heading5TextStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING5, heading5TextStyle);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING6, heading5TextStyle);*/
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING5, customStyle['h5']);
    body.setHeadingAttributes(DocumentApp.ParagraphHeading.HEADING6, customStyle['h6']);
    // End. Set up heading 5(6) (named style type HEADING_5) attributes

    // Set up footnotes attributes
    /* const footnoteStyle = {};
    footnoteStyle[DocumentApp.Attribute.FONT_FAMILY] = config_fontFamily;
    footnoteStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
    footnoteStyle[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    footnoteStyle[DocumentApp.Attribute.SPACING_AFTER] = 10;
    footnoteStyle[DocumentApp.Attribute.LINE_SPACING] = 1.15;*/
    const footnotes = doc.getFootnotes();
    footnotes.forEach(function (item) {
      item.getFootnoteContents().getParagraphs().forEach(function (item) {
        item.setAttributes(customStyle['footnote']);
      });
    });
    // End. Set up footnotes attributes

    // Set up lists attributes part 1
    Logger.log('formatListsPart1');
    formatListsPart1(false, requests, body, document, documentId);

    // Set up 20pt after tables
    setSpaceAfterTables20pt(body);

    setSpaceBeforeAfterInTable(body);

    doc.saveAndClose();

    // Set up lists attributes part 2
    //formatListsPart2(false, requests, document, documentId);

    //Logger.log('formatHeader');
    const resultHeader = formatHeader();
    if (resultHeader.status == 'error') {
      ui.alert(resultHeader.message);
    }

    //Logger.log('formatFooter');
    const result = formatFooter();
    if (result.status == 'error') {
      ui.alert(result.message);
    }


    document = Docs.Documents.get(documentId);
    const bodyElements = document.body.content;

    const allFootnotes = document.footnotes;
    for (let footnoteId in allFootnotes) {
      allFootnotes[footnoteId].content.forEach(function (content) {
        content.paragraph.elements.forEach(function (item) {
          // Remove underline from hyperlinks
          helpRemoveUnderlineFromHyperlinks(requests, item, footnoteId);
        });
      });
    }

    const arrayH1H2 = [];
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

        const namedStyleType = bodyElements[i].paragraph.paragraphStyle.namedStyleType;
        const elements = bodyElements[i].paragraph.elements;
        const lastElement = elements.length - 1;

        if (['HEADING_1', 'HEADING_2', 'HEADING_3', 'HEADING_4', 'HEADING_5', 'HEADING_6'].includes(namedStyleType)) {
          // If paragraph has HEADING_1, HEADING_2... named style, we push it in array arrayH1H2
          arrayH1H2.push({
            style: namedStyleType,
            startIndex: elements[0].startIndex,
            endIndex: elements[lastElement].endIndex
          });
        } else if (namedStyleType == 'NORMAL_TEXT') {
          // If paragraph has NORMAL_TEXT named style

          // Check paragraph's style
          const spaceBelow = bodyElements[i].paragraph.paragraphStyle.spaceBelow;
          const spaceAbove = bodyElements[i].paragraph.paragraphStyle.spaceAbove;

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

    // Bottom border of H2/H3
    const transparentBorder = { width: { magnitude: 0, unit: 'PT' }, padding: { magnitude: 0, unit: 'PT' }, dashStyle: 'SOLID' };
    if (styles[ACTIVE_STYLE]['headingBorderBottom'].hasOwnProperty('h2')) {
      h1H2styles.paragraphStyle_HEADING_2.borderBottom = styles[ACTIVE_STYLE]['headingBorderBottom']['h2'];
    } else {
      h1H2styles.paragraphStyle_HEADING_2.borderBottom = transparentBorder;
    }
    if (styles[ACTIVE_STYLE]['headingBorderBottom'].hasOwnProperty('h3')) {
      h1H2styles.paragraphStyle_HEADING_3.borderBottom = styles[ACTIVE_STYLE]['headingBorderBottom']['h3'];
    } else {
      h1H2styles.paragraphStyle_HEADING_3.borderBottom = transparentBorder;
    }
    // End. Bottom border of H2/H3

    addTextStyleHeading();

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


    const l = requests.length - 1;
    for (let i = l; i >= 0; i--) {
      if (requests[i].updateParagraphStyle) {
        if (!requests[i].updateParagraphStyle.paragraphStyle) {
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
    return {status: 'ok'};
  } catch (error) {
    Logger.log('Error in defaultStyleReport: ' + error);
    ui.alert('Error in defaultStyleReport: ' + error);
    return {status: 'error'};
  }
}

// Check fontSize and fontFamily of element of paragraph
// defaultStyleReport and formatListsPart2 use the function
function checkElementOfParagraph(requests, item, updateParagraphStyle) {

  const config_fontFamily = styles[ACTIVE_STYLE]['fontFamily'];

  let normalTextForegroundColor, likeH1textColor = false, wrongFontSize = false;
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
