function defaultMargins() {
  let ui = DocumentApp.getUi();
  try {
    let doc = DocumentApp.getActiveDocument();
    let body = doc.getBody();

    let style = {};
    style[DocumentApp.Attribute.MARGIN_TOP] = '56.69291338582678';
    style[DocumentApp.Attribute.MARGIN_BOTTOM] = '56.69291338582678';
    style[DocumentApp.Attribute.MARGIN_LEFT] = '56.69291338582678';
    style[DocumentApp.Attribute.MARGIN_RIGHT] = '56.69291338582678';
    body.setAttributes(style);
  }
  catch (error) {
    ui.alert('Error in defaultMargins: ' + error);
  }
}
