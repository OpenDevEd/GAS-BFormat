const paragraphStyle_FIGURE_HEADING_5 = {
  namedStyleType: styles[ACTIVE_STYLE]['figureHeading'],
  spaceAbove: { magnitude: 10, unit: 'PT' },
  spaceBelow: { magnitude: 6, unit: 'PT' },
  alignment: 'START'
};

const paragraphStyle_FIGURE_CONTENT = {
  namedStyleType: 'NORMAL_TEXT',
  spaceAbove: { magnitude: 10, unit: 'PT' },
  spaceBelow: { magnitude: 6, unit: 'PT' },
  alignment: 'CENTER'
};

const textStyle_FIGURE_PART_1 = {
  foregroundColor: {
    color: {
      rgbColor: hexToRGB(styles[ACTIVE_STYLE]['customStyle']['h6']['FOREGROUND_COLOR'])
    }
  },
  fontSize: {
    magnitude: styles[ACTIVE_STYLE]['customStyle']['h6']['FONT_SIZE'],
    unit: 'PT'
  },
  bold: true,
  italic: false,
  weightedFontFamily: {
    fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
    weight: 400
  }
};

const textStyle_FIGURE_PART_2 = {
  foregroundColor: {
    color: {
      rgbColor: hexToRGB(styles[ACTIVE_STYLE]['customStyle']['h6']['FOREGROUND_COLOR'])
    }
  },
  fontSize: {
    magnitude: styles[ACTIVE_STYLE]['customStyle']['h6']['FONT_SIZE'],
    unit: 'PT'
  },
  bold: false,
  italic: true,
  weightedFontFamily: {
    fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
    weight: 400
  }
};

const textStyle_FIGURE_CONTENT = {
  foregroundColor: {
    color: {
      rgbColor: {red:0, green:0, blue:0}
    }
  },  
  fontSize: {
    magnitude: 12,
    unit: 'PT'
  },
  bold: false,
  italic: false,
  weightedFontFamily: {
    fontFamily: styles[ACTIVE_STYLE]['fontFamily'],
    weight: 400
  }
};

// Inserts text
// Figure 1...
// [figure content or Image here]
function insertFigure1() {
  const ui = DocumentApp.getUi();
  try {
    const textFigure1H5 = 'Figure 1. Image Caption Title if needed. I am an image caption description. I tell people what the image is about. I also acknowledge the image source.\n';
    const textFigureContentNormal = '[figure content or Image here]';

    const lenTextFigure1H5 = textFigure1H5.length;
    const lenTextFigureContentNormal = textFigureContentNormal.length;


    const doc = DocumentApp.getActiveDocument();
    const documentId = doc.getId();

    const cursorPosition = detectCursorPosition(doc, documentId);
    if (cursorPosition.status == 'error') {
      ui.alert(cursorPosition.message);
      return 0;
    }

    const insertStartIndex = cursorPosition.endIndex;

    const requests = [];

    requests.push({
      insertText: {
        text: textFigure1H5,
        location: {
          index: insertStartIndex
        }
      }
    },
      {
        insertText: {
          text: textFigureContentNormal,
          location: {
            index: insertStartIndex + lenTextFigure1H5
          }
        }
      },
      {
        updateParagraphStyle: {
          paragraphStyle: paragraphStyle_FIGURE_HEADING_5,
          range: {
            startIndex: insertStartIndex,
            endIndex: insertStartIndex + lenTextFigure1H5
          },
          fields: formFieldsString(paragraphStyle_FIGURE_HEADING_5)
        }
      },
      {
        updateParagraphStyle: {
          paragraphStyle: paragraphStyle_FIGURE_CONTENT,
          range: {
            startIndex: insertStartIndex + lenTextFigure1H5,
            endIndex: insertStartIndex + lenTextFigure1H5 + lenTextFigureContentNormal
          },
          fields: formFieldsString(paragraphStyle_FIGURE_CONTENT)
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
            endIndex: insertStartIndex + lenTextFigure1H5
          },
          text_style: textStyle_FIGURE_PART_2,
          fields: formFieldsString(textStyle_FIGURE_PART_2)
        }
      },
      {
        updateTextStyle: {
          range: {
            startIndex: insertStartIndex + lenTextFigure1H5,
            endIndex: insertStartIndex + lenTextFigure1H5 + lenTextFigureContentNormal
          },
          text_style: textStyle_FIGURE_CONTENT,
          fields: formFieldsString(textStyle_FIGURE_CONTENT)
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
    ui.alert('Error in insertFigure1. ' + error);
  }
}

// Inserts 1 row 2 columns table
// [figure content or Image here]  Figure 2...
function insertFigure2() {
  const ui = DocumentApp.getUi();
  try {
    const textFigure2H5 = 'Figure 2. Image Caption Title if needed. I am an image caption description. I tell people what the image is about. I also acknowledge the image source.';
    const textFigureContentNormal = '[figure content or Image here]';

    const lenTextFigure2H5 = textFigure2H5.length;
    const lenTextFigureContentNormal = textFigureContentNormal.length;


    const doc = DocumentApp.getActiveDocument();
    const documentId = doc.getId();

    const cursorPosition = detectCursorPosition(doc, documentId);
    if (cursorPosition.status == 'error') {
      ui.alert(cursorPosition.message);
      return 0;
    }

    const insertStartIndex = cursorPosition.endIndex;

    const requests = [];

    requests.push(
      {
        insertTable: {
          rows: 1,
          columns: 2,
          location: { index: insertStartIndex }
        }
      },
      {
        updateTableCellStyle: {
          tableRange: {
            tableCellLocation: {
              tableStartLocation: {
                index: insertStartIndex + 1
              },
            },
            rowSpan: 1,
            columnSpan: 2
          },

          tableCellStyle: {
            borderTop: tableStyle_TRANSPERENT_BORDER,
            borderBottom: tableStyle_TRANSPERENT_BORDER,
            borderLeft: tableStyle_TRANSPERENT_BORDER,
            borderRight: tableStyle_TRANSPERENT_BORDER,
          },
          fields: 'borderTop,borderBottom,borderLeft,borderRight'
        }
      },
      {
        insertText: {
          text: textFigure2H5,
          location: {
            index: insertStartIndex + 6
          }
        }
      },
      {
        insertText: {
          text: textFigureContentNormal,
          location: {
            index: insertStartIndex + 4
          }
        }
      },


      {
        updateParagraphStyle: {
          paragraphStyle: paragraphStyle_FIGURE_HEADING_5,
          range: {
            startIndex: insertStartIndex + 6 + lenTextFigureContentNormal,
            endIndex: insertStartIndex + 6 + lenTextFigureContentNormal + lenTextFigure2H5
          },
          fields: formFieldsString(paragraphStyle_FIGURE_HEADING_5)
        }
      },
      {
        updateParagraphStyle: {
          paragraphStyle: paragraphStyle_FIGURE_CONTENT,
          range: {
            startIndex: insertStartIndex + 4,
            endIndex: insertStartIndex + 4 + lenTextFigureContentNormal
          },
          fields: formFieldsString(paragraphStyle_FIGURE_CONTENT)
        }
      },
      {
        updateTextStyle: {
          range: {
            startIndex: insertStartIndex + 6 + lenTextFigureContentNormal,
            endIndex: insertStartIndex + 6 + lenTextFigureContentNormal + 9
          },
          text_style: textStyle_FIGURE_PART_1,
          fields: formFieldsString(textStyle_FIGURE_PART_1)
        }
      },
      {
        updateTextStyle: {
          range: {
            startIndex: insertStartIndex + 6 + lenTextFigureContentNormal + 9,
            endIndex: insertStartIndex + 6 + lenTextFigureContentNormal + lenTextFigure2H5
          },
          text_style: textStyle_FIGURE_PART_2,
          fields: formFieldsString(textStyle_FIGURE_PART_2)
        }
      },
      {
        updateTextStyle: {
          range: {
            startIndex: insertStartIndex + 4,
            endIndex: insertStartIndex + 4 + lenTextFigureContentNormal
          },
          text_style: textStyle_FIGURE_CONTENT,
          fields: formFieldsString(textStyle_FIGURE_CONTENT)
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
    ui.alert('Error in insertFigure2. ' + error);
  }
}
