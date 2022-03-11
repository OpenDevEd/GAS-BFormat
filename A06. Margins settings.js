function defaultMargins() {
  let ui = DocumentApp.getUi();
  try {
    let doc = DocumentApp.getActiveDocument();
    let body = doc.getBody();

    let style = {};
    // 2cm = 56.69291338582678 pt
    const cmTOpt = 56.69291338582678 / 2;
    style[DocumentApp.Attribute.MARGIN_TOP] = styles[getThisDocStyle()]['MARGIN_TOP_cm'] * cmTOpt;
    style[DocumentApp.Attribute.MARGIN_BOTTOM] = styles[getThisDocStyle()]['MARGIN_BOTTOM_cm']  * cmTOpt;
    style[DocumentApp.Attribute.MARGIN_LEFT] = 56.69291338582678;
    style[DocumentApp.Attribute.MARGIN_RIGHT] = 56.69291338582678;
    body.setAttributes(style);
  }
  catch (error) {
    ui.alert('Error in defaultMargins: ' + error);
  }
}
