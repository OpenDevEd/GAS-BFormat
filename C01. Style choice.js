// Updates thisDocStyle document property, updates rgbColor and fontFamily of all style objects that contains color and font family
function useStyle(styleName) {
  // setDocumentPropertyString('thisDocStyle', styleName);

  ACTIVE_STYLE = styleName;
  Logger.log('useStyle=' + ACTIVE_STYLE);

  const rgbColor = hexToRGB(styles[ACTIVE_STYLE]['main_heading_font_color']);
  const h6RgbColor = hexToRGB(styles[ACTIVE_STYLE]['customStyle']['h6']['FOREGROUND_COLOR']);
  const fontFamily = styles[ACTIVE_STYLE]['fontFamily'];

  // Updates objects in file A01. Default style Report
  h1H2styles.textStyle_HEADING_1.foregroundColor.color.rgbColor = rgbColor;
  h1H2styles.textStyle_HEADING_1.weightedFontFamily.fontFamily = fontFamily;
  h1H2styles.textStyle_HEADING_2.weightedFontFamily.fontFamily = fontFamily;

  h1H2styles.textStyle_HEADING_1.fontSize.magnitude = styles[ACTIVE_STYLE]['customStyle']['h1']['FONT_SIZE'];
  h1H2styles.textStyle_HEADING_2.fontSize.magnitude = styles[ACTIVE_STYLE]['customStyle']['h2']['FONT_SIZE'];

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
  textStyle_TABLE_HEADING_PART_1.foregroundColor.color.rgbColor = h6RgbColor;
  textStyle_TABLE_HEADING_PART_2.weightedFontFamily.fontFamily = fontFamily;
  textStyle_TABLE_HEADING_PART_2.foregroundColor.color.rgbColor = h6RgbColor;

  textStyle_TABLE_HEADING_PART_1.fontSize.magnitude = styles[ACTIVE_STYLE]['customStyle']['h6']['FONT_SIZE'];
  textStyle_TABLE_HEADING_PART_2.fontSize.magnitude = styles[ACTIVE_STYLE]['customStyle']['h6']['FONT_SIZE'];

  paragraphStyle_TABLE.spaceAbove = styles[ACTIVE_STYLE]['paragraphSpacesInCell'];
  paragraphStyle_TABLE.spaceBelow = styles[ACTIVE_STYLE]['paragraphSpacesInCell'];

  // Updates objects in file A05. Figure insertion
  textStyle_FIGURE_PART_1.foregroundColor.color.rgbColor = rgbColor;
  textStyle_FIGURE_PART_1.weightedFontFamily.fontFamily = fontFamily;
  textStyle_FIGURE_PART_2.foregroundColor.color.rgbColor = rgbColor;
  textStyle_FIGURE_PART_2.weightedFontFamily.fontFamily = fontFamily;
  textStyle_FIGURE_CONTENT.foregroundColor.color.rgbColor = rgbColor;
  textStyle_FIGURE_CONTENT.weightedFontFamily.fontFamily = fontFamily;

  textStyle_FIGURE_PART_1.fontSize.magnitude = styles[ACTIVE_STYLE]['customStyle']['h6']['FONT_SIZE'];
  textStyle_FIGURE_PART_2.fontSize.magnitude = styles[ACTIVE_STYLE]['customStyle']['h6']['FONT_SIZE'];

  // Updates objects in file A08. Format header
  //paragraphStyle_HEADING_SEC.borderBottom.color.color.rgbColor = rgbColor;
  if (styles[ACTIVE_STYLE]["headerBottom"] === true) {
    paragraphStyle_HEADING_SEC.borderBottom = paragraphStyle_HEADING_SEC_PLUS_BOTTOM;
  }
  textStyle_HEADING_SEC.weightedFontFamily.fontFamily = fontFamily;

  // Updates objects in file A09. Format footer
  textStyle_FOOTER_SEC.foregroundColor.color.rgbColor = rgbColor;
  textStyle_FOOTER_SEC.weightedFontFamily.fontFamily = fontFamily;

  // Updates objects in file B04. Block quote style
  textStyle_EXTRACTED_QUOTE_1.weightedFontFamily.fontFamily = fontFamily;
  textStyle_EXTRACTED_QUOTE_2.weightedFontFamily.fontFamily = fontFamily;

  // Updates object in file B05. Right border
  paragraphStyle_LEFT_BORDER.borderLeft.color.color.rgbColor = rgbColor;

  const { status } = defaultStyleReport();
  if (status === 'ok') {
    setDocumentPropertyString('thisDocStyle', styleName);
    onOpen();
    const updatedMenuStructure = universal_bFormat_menu(null, 'data');
    return {
      needUpdate: true,
      updatedMenuData: updatedMenuStructure
    };
  }
}