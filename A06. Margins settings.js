function defaultMargins() {
  const ui = DocumentApp.getUi();
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    const style = {};
    // 2cm = 56.69291338582678 pt
    style[DocumentApp.Attribute.MARGIN_TOP] = styles[ACTIVE_STYLE]['MARGIN_TOP_cm'] * cmTOpt;
    style[DocumentApp.Attribute.MARGIN_BOTTOM] = styles[ACTIVE_STYLE]['MARGIN_BOTTOM_cm']  * cmTOpt;
    style[DocumentApp.Attribute.MARGIN_LEFT] = styles[ACTIVE_STYLE]['MARGIN_LEFT_cm'] * cmTOpt;
    style[DocumentApp.Attribute.MARGIN_RIGHT] = styles[ACTIVE_STYLE]['MARGIN_RIGHT_cm'] * cmTOpt;
    body.setAttributes(style);
  }
  catch (error) {
    ui.alert('Error in defaultMargins: ' + error);
  }
}
