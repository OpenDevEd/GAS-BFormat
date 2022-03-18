const styles = {
  "report_default":
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
  },
  "report_opendeved":
  {
    "name": "Report (OpenDevEd)",
    "default_everybody": false,
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
  const activeUser = Session.getEffectiveUser().getEmail();
  for (let styleName in styles) {
    if (styles[styleName]['default_for'] && activeUser.search(new RegExp(styles[styleName]['default_for'], 'i')) != -1) {
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

// The variable contains style of current doc
// Initially, it is default style but 
// (1) function updateStyle changes it to style stored in DocumentProperties thisDocStyle
// (2) function useStyle changes it to style selected by user
let ACTIVE_STYLE = getDefaultStyle();

// Changes a value of ACTIVE_STYLE to thisDocStyle doc property if thisDocStyle doc property is set for an active document
function updateStyle() {
  Logger.log(' updateStyle()');
  try {
    const thisDocStyle = getDocumentPropertyString('thisDocStyle');
    if (thisDocStyle != null) {
      ACTIVE_STYLE = thisDocStyle;
    }
  }
  catch (error) {
    Logger.log(error);
  }
}
updateStyle();

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