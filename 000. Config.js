const styles = {
  "report_default":
  {
    "name": "Report (default)",
    "fontFamily": "Open Sans",
    "default_everybody": true,
    "MARGIN_TOP_cm": 2.5,
    "MARGIN_BOTTOM_cm": 2.5,
    "MARGIN_HEADER_cm": 1.0,
    "MARGIN_FOOTER_cm": 1.0,
    "main_heading_font_color": "#6aa84f",
    //"main_heading_font_color": "#FF0000",
  },
  "report_opendeved":
  {
    "name": "Report (OpenDevEd)",
    "default_everbody": false,
    "default_for": "opendeved.net",
    "fontFamily": "Ubuntu",
    "MARGIN_TOP_cm": 2.5,
    "MARGIN_BOTTOM_cm": 2.5,
    "MARGIN_HEADER_cm": 1.0,
    "MARGIN_FOOTER_cm": 1.0,
    "main_heading_font_color": "#FF5C00",
    // "main_heading_font_color": "#000000",
    // "main_heading_font_color": "#EF8456",
  },
  "report_edtechhub":
  {
    "name": "Report (EdTech Hub)",
    "default_everbody": false,
    "default_for": "edtechhub.org",
    "fontFamily": "Montserrat",
    "MARGIN_TOP_cm": 2.0,
    "MARGIN_BOTTOM_cm": 2.0,
    "MARGIN_HEADER_cm": 1.0,
    "MARGIN_FOOTER_cm": 1.0,
    "main_heading_font_color": "#FF5C00",
  }
};



function getDefaultEverybodyStyle() {
  for (let styleName in styles) {
    if (styles[styleName]['default_everybody'] === true) {
      return styleName;
    }
  }
}


function getDefaultStyle() {
  const activeUser = Session.getEffectiveUser().getEmail();
  for (let styleName in styles) {
    if (styles[styleName]['default_for'] && activeUser.search(new RegExp(styles[styleName]['default_for'], 'i')) != -1) {
      return styleName;
    }
  }

  return getDefaultEverybodyStyle();
}




let activeStyle = getDefaultStyle();


function updateStyle() {
  Logger.log(' updateStyle()');
  try {
    const thisDocStyle = getDocumentPropertyString('thisDocStyle');
    if (thisDocStyle != null) {
      activeStyle = thisDocStyle;
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

function setDocumentPropertyString(property_name, value) {
  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(property_name, value);
}

function getDocumentPropertyString(property_name) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const value = documentProperties.getProperty(property_name);
  return value;
}

function getThisDocStyle() {
  Logger.log(activeStyle);
  return activeStyle;
}
