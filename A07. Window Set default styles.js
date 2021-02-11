function setDefaultStylesManually() {
  var htmlOutput = HtmlService
    .createHtmlOutput(`<p>To create default styles, please open <a target="_blank" href="https://docs.google.com/document/d/1LUD9DWi1rvi5SoKV-mpanmjPD4vXRLX2ACM-uRdk9YQ/edit">TEMPLATE</a>. In the template, do the following:</p>
    <ul>
    <li>Format > Paragraph Styles > Options > Save as my default styles</li>
    <li>You can then close the template</li>
    </ul>
    
    <img src="https://drive.google.com/uc?export=view&id=1qeB4MPI3QPMYoGgHQ91bl_zJHGNaWI5H">`)
    .setWidth(620)
    .setHeight(580);
  DocumentApp.getUi().showModalDialog(htmlOutput, 'BFormat');


}
