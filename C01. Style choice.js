// Updates thisDocStyle document property, updates rgbColor and fontFamily of all style objects that contains color and font family
function useStyle(styleName) {
  setDocumentPropertyString('thisDocStyle', styleName);

  ACTIVE_STYLE = styleName;
  Logger.log('useStyle=' + ACTIVE_STYLE);

  const rgbColor = hexToRGB(styles[ACTIVE_STYLE]['main_heading_font_color']);
  const fontFamily = styles[ACTIVE_STYLE]['fontFamily'];

  // Updates objects in file A01. Default style Report
  h1H2styles.textStyle_HEADING_1.foregroundColor.color.rgbColor = rgbColor;
  h1H2styles.textStyle_HEADING_1.weightedFontFamily.fontFamily = fontFamily;
  h1H2styles.textStyle_HEADING_2.weightedFontFamily.fontFamily = fontFamily;

  // Updates objects in file A02. Quote insertion
  paragraphStyle_QUOTE_1.borderTop.color.color.rgbColor = rgbColor;
  paragraphStyle_QUOTE_2.borderBottom.color.color.rgbColor = rgbColor;
  textStyle_QUOTE.foregroundColor.color.rgbColor = rgbColor;
  textStyle_QUOTE.weightedFontFamily.fontFamily = fontFamily;

  // Updates objects in file A03. Table insertion
  tableStyles.textStyle_ITEM_CELL.weightedFontFamily.fontFamily = fontFamily;
  tableStyles.textStyle_TOPIC_COLUMN_CELL.weightedFontFamily.fontFamily = fontFamily;
  tableStyle_ORANGE_BORDER.color.color.rgbColor = rgbColor;
  textStyle_TABLE_HEADING_PART_1.weightedFontFamily.fontFamily = fontFamily;
  textStyle_TABLE_HEADING_PART_1.foregroundColor.color.rgbColor = rgbColor;
  textStyle_TABLE_HEADING_PART_2.weightedFontFamily.fontFamily = fontFamily;
  textStyle_TABLE_HEADING_PART_2.foregroundColor.color.rgbColor = rgbColor;

  // Updates objects in file A05. Figure insertion
  textStyle_FIGURE_PART_1.foregroundColor.color.rgbColor = rgbColor;
  textStyle_FIGURE_PART_1.weightedFontFamily.fontFamily = fontFamily;
  textStyle_FIGURE_PART_2.foregroundColor.color.rgbColor = rgbColor;
  textStyle_FIGURE_PART_2.weightedFontFamily.fontFamily = fontFamily;
  textStyle_FIGURE_CONTENT.foregroundColor.color.rgbColor = rgbColor;
  textStyle_FIGURE_CONTENT.weightedFontFamily.fontFamily = fontFamily;

  // Updates objects in file A08. Format header
  paragraphStyle_HEADING_SEC.borderBottom.color.color.rgbColor = rgbColor;
  textStyle_HEADING_SEC.weightedFontFamily.fontFamily = fontFamily;

  // Updates objects in file A09. Format footer
  textStyle_FOOTER_SEC.foregroundColor.color.rgbColor = rgbColor;
  textStyle_FOOTER_SEC.weightedFontFamily.fontFamily = fontFamily;

  // Updates objects in file B04. Block quote style
  textStyle_EXTRACTED_QUOTE_1.weightedFontFamily.fontFamily = fontFamily;
  textStyle_EXTRACTED_QUOTE_2.weightedFontFamily.fontFamily = fontFamily;

  // Updates object in file B05. Right border
  paragraphStyle_LEFT_BORDER.borderLeft.color.color.rgbColor = rgbColor;

  defaultStyleReport();
  onOpen();
}