const styles = {
  /* "report_default":
  {
    "name": "Report (default)",
    "fontFamily": "Open Sans",
    "default_everybody": true,
    "MARGIN_TOP_cm": 2.5,
    "MARGIN_BOTTOM_cm": 2.5,
    "MARGIN_LEFT_cm": 2.0,
    "MARGIN_RIGHT_cm": 2.0,
    "MARGIN_HEADER_cm": 1.0,
    "MARGIN_FOOTER_cm": 1.0,
    "main_heading_font_color": "#6aa84f",
    "title_position": "header",
    "header_text": "Title of document as shown on cover page",
    "footer_text": "Your organisation",
    "header_font_size": 11,
    "footer_font_size": 11,
  }, */
  "report_opendeved":
  {
    "name": "Report (OpenDevEd)",
    "default_everybody": true,
    "default_for": "opendeved.net",
    "fontFamily": "Ubuntu",
    "pageWidth_cm": 21,
    "pageHeight_cm": 29.7,
    "MARGIN_TOP_cm": 2.5,
    "MARGIN_BOTTOM_cm": 2.5,
    "MARGIN_LEFT_cm": 2.0,
    "MARGIN_RIGHT_cm": 2.0,
    "MARGIN_HEADER_cm": 1.0,
    "MARGIN_FOOTER_cm": 1.0,
    "main_heading_font_color": "#E68225",
    "title_position": "header",
    "FOOTER": true,
    "header_text": "Title of document as shown on cover page",
    "footer_text": "OpenDevEd",
    "header_font_size": 11,
    "footer_font_size": 11,
    "headerBottom": true,
    "figureHeading": "HEADING_6",
    "glyphType": "SQUARE_BULLET",
    "paragraphSpacesInCell":10,
    "customStyle": {
      h1: { FONT_SIZE: 21, FOREGROUND_COLOR: '#E68225'},
      h2: { FONT_SIZE: 16 },
      h6: { FONT_SIZE: 11, SPACING_BEFORE: 0, SPACING_AFTER: 10, LINE_SPACING: 1.15, FOREGROUND_COLOR: '#E68225', ITALIC: true },
    },
    "headingBorderBottom": {
      h2: { width: { magnitude: 1, unit: 'PT' }, padding: { magnitude: 2, unit: 'PT' }, dashStyle: 'SOLID' }
    }
  },
  "report_edtechhub":
  {
    "name": "Report (EdTech Hub)",
    "default_everybody": false,
    "default_for": "edtechhub.org",
    "fontFamily": "Montserrat",
    "pageWidth_cm": 21,
    "pageHeight_cm": 29.7,
    "MARGIN_TOP_cm": 0,
    "MARGIN_BOTTOM_cm": 0,
    "MARGIN_LEFT_cm": 2.54,
    "MARGIN_RIGHT_cm": 2.54,
    "MARGIN_HEADER_cm": 1.27,
    "MARGIN_FOOTER_cm": 1.27,
    "main_heading_font_color": "#FF5C00",
    "title_position": "footer",
    "FOOTER": false,
    "header_text": "EdTech Hub",
    "footer_text": "Title of document as shown on cover page",
    "header_font_size": 10,
    "footer_font_size": 10,
    "headerBottom": true,
    "figureHeading": "HEADING_6",
    "glyphType": "SQUARE_BULLET",
    "paragraphSpacesInCell":10,
    "customStyle": {
      h1: { FONT_SIZE: 21, SPACING_BEFORE: 14, SPACING_AFTER: 10, LINE_SPACING: 1.15, INDENT_FIRST_LINE: 0, FOREGROUND_COLOR: '#FF5C00', BOLD: true, ITALIC: false},
      h2: { FONT_SIZE: 16, SPACING_BEFORE: 14, SPACING_AFTER: 8, LINE_SPACING: 1.15, FOREGROUND_COLOR: '#000000', BOLD: true, ITALIC: false},
      h3: { FONT_SIZE: 14, SPACING_BEFORE: 0, SPACING_AFTER: 10, LINE_SPACING: 1.15, FOREGROUND_COLOR: '#000000', BOLD: true, ITALIC: false},
      h4: { FONT_SIZE: 12, SPACING_BEFORE: 0, SPACING_AFTER: 10, LINE_SPACING: 1.15, FOREGROUND_COLOR: '#000000', BOLD: true, ITALIC: false},
      h5: { FONT_SIZE: 11, SPACING_BEFORE: 0, SPACING_AFTER: 10, LINE_SPACING: 1.15, FOREGROUND_COLOR: '#000000', ITALIC: true },
      h6: { FONT_SIZE: 11, SPACING_BEFORE: 0, SPACING_AFTER: 10, LINE_SPACING: 1.15, FOREGROUND_COLOR: '#FF5C00', ITALIC: true },
    },
    "headingBorderBottom": {
      h2: { width: { magnitude: 1, unit: 'PT' }, padding: { magnitude: 2, unit: 'PT' }, dashStyle: 'SOLID' }
    }
  },
  "report_EdTech_Fellowship":
  {
    "name": "Report (EdTech Fellowship)",
    "default_everybody": false,
    //"default_for": "edtechhub.org",
    "fontFamily": "Open Sans",
    "pageWidth_cm": 21,
    "pageHeight_cm": 29.7,
    "MARGIN_TOP_cm": 3.5,
    "MARGIN_BOTTOM_cm": 3.0,
    "MARGIN_LEFT_cm": 2.5,
    "MARGIN_RIGHT_cm": 2.5,
    "MARGIN_HEADER_cm": 1.27,
    "MARGIN_FOOTER_cm": 0.9,
    "main_heading_font_color": "#133844",
    "title_position": "footer",
    "FOOTER": true,
    "header_text": "EdTech Fellowship",
    "footer_text": "The HP Cambridge EdTech Fellowship",
    "header_font_size": 11,
    "headerBottom": false,
    "footer_font_size": 11,
    "figureHeading": "HEADING_6",
    "glyphType": "BULLET",
    "paragraphSpacesInCell":5,
    "customStyle": {
      normalText: { FONT_SIZE: 12, SPACING_BEFORE: 0, SPACING_AFTER: 10, LINE_SPACING: 1.15, FOREGROUND_COLOR: '#212529' },
      h1: { FONT_SIZE: 36, SPACING_BEFORE: 30, SPACING_AFTER: 20, LINE_SPACING: 1, FOREGROUND_COLOR: '#133844' },
      h2: { FONT_SIZE: 24, SPACING_BEFORE: 24, SPACING_AFTER: 0, LINE_SPACING: 1.4, FOREGROUND_COLOR: '#3c1366' },
      h3: { FONT_SIZE: 21, SPACING_BEFORE: 24, SPACING_AFTER: 10, LINE_SPACING: 1.15, FOREGROUND_COLOR: '#00bdb6' },
      h4: { FONT_SIZE: 18, SPACING_BEFORE: 24, SPACING_AFTER: 10, LINE_SPACING: 1.15, FOREGROUND_COLOR: '#212529' },
      h5: { FONT_SIZE: 16, SPACING_BEFORE: 24, SPACING_AFTER: 10, LINE_SPACING: 1.15, FOREGROUND_COLOR: '#212529', BOLD: true, ITALIC: true },
      h6: { FONT_SIZE: 13, SPACING_BEFORE: 24, SPACING_AFTER: 10, LINE_SPACING: 1.15, FOREGROUND_COLOR: '#8128e7', BOLD: false, ITALIC: true },
    },
    "headingBorderBottom": {
      h3: {
        width: { magnitude: 1, unit: 'PT' }, padding: { magnitude: 2, unit: 'PT' }, dashStyle: 'SOLID', color: { color: { rgbColor: hexToRGB('#00bdb6') } }
      }
    }
  }
};


// Gets default style based on user's domain
function getDefaultStyle() {
  const activeUser = Session.getActiveUser().getEmail().toLowerCase();
  for (let styleName in styles) {
    if (styles[styleName]['default_for'] && activeUser.search(styles[styleName]['default_for']) != -1) {
      return styleName;
    }
  }
  // If user's domain isn't presented in styles object, find style that is suitable for everybody
  for (let styleName in styles) {
    if (styles[styleName]['default_everybody'] === true) {
      return styleName;
    }
  }
}

// Detects style of current doc
function getThisDocumentStyle(tryToRetrieveProperties) {
  const DEFAULT_DOC_STYLE = getDefaultStyle();
  const resultObj = {
    marker: '☑️',
    style: DEFAULT_DOC_STYLE,
    domainBasedStyle: DEFAULT_DOC_STYLE,
    menuText: styles[DEFAULT_DOC_STYLE]['name']
  };
  if (tryToRetrieveProperties === true) {
    try {
      const docProperties = PropertiesService.getDocumentProperties();
      const docStyle = docProperties.getProperty('thisDocStyle');
      if (docStyle != null && styles.hasOwnProperty(docStyle)) {
        resultObj['style'] = docStyle;
        resultObj['menuText'] = styles[docStyle]['name'];
        resultObj['marker'] = '✅';
      }
    }
    catch (error) {
      Logger.log('Needs to activate!!! ' + error);
    }
  }

  return resultObj;
}

// The variable contains style of current doc
// function useStyle can change it to style selected by user
let ACTIVE_STYLE = getThisDocumentStyle(true).style;

function report_default() {
  return useStyle('report_default');
}

function report_opendeved() {
  return useStyle('report_opendeved');
}

function report_edtechhub() {
  return useStyle('report_edtechhub');
}

function report_EdTech_Fellowship() {
  return useStyle('report_EdTech_Fellowship');
}

// Sets document property
function setDocumentPropertyString(property_name, value) {
  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(property_name, value);
}

// Retrieves document property
function getDocumentPropertyString(property_name) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const value = documentProperties.getProperty(property_name);
  return value;
}