function setDefaultStylesManually() {
  const htmlOutput = HtmlService
    .createHtmlOutput(`<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
  </head>
  <body><p>After applying your preferred style, use the menu options below to set your preferred style as default:</p>
    <ul>
    <li>Format > Paragraph Styles > Options > Save as my default styles</li>
    <li>You can then close the template</li>
    </ul>
    
    <img src="https://drive.google.com/uc?export=view&id=1qeB4MPI3QPMYoGgHQ91bl_zJHGNaWI5H">  
    <p>If you don't see the image above, click <a target="_blank" href="https://drive.google.com/uc?export=view&id=1qeB4MPI3QPMYoGgHQ91bl_zJHGNaWI5H">here</a>.</p>
    </body>
</html>`)
    .setWidth(620)
    .setHeight(580);
  DocumentApp.getUi().showModalDialog(htmlOutput, 'bFormat');
}