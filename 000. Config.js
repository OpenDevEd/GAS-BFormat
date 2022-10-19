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
    "MARGIN_TOP_cm": 2.5,
    "MARGIN_BOTTOM_cm": 2.5,
    "MARGIN_LEFT_cm": 2.0,
    "MARGIN_RIGHT_cm": 2.0,
    "MARGIN_HEADER_cm": 1.0,
    "MARGIN_FOOTER_cm": 1.0,
    "main_heading_font_color": "#FF5C00",
    "title_position": "header",
    "header_text": "Title of document as shown on cover page",
    "footer_text": "OpenDevEd",
    "header_font_size": 11,
    "footer_font_size": 11,
  },
  "report_edtechhub":
  {
    "name": "Report (EdTech Hub)",
    "default_everybody": false,
    "default_for": "edtechhub.org",
    "fontFamily": "Montserrat",
    "MARGIN_TOP_cm": 2.0,
    "MARGIN_BOTTOM_cm": 2.0,
    "MARGIN_LEFT_cm": 2.0,
    "MARGIN_RIGHT_cm": 2.0,
    "MARGIN_HEADER_cm": 1.0,
    "MARGIN_FOOTER_cm": 1.0,
    "main_heading_font_color": "#FF5C00",
    "title_position": "footer",
    "header_text": "EdTech Hub",
    "footer_text": "Title of document as shown on cover page",
    "header_font_size": 10,
    "footer_font_size": 10,
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
  useStyle('report_default');
}

function report_opendeved() {
  useStyle('report_opendeved');
}

function report_edtechhub() {
  useStyle('report_edtechhub');
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