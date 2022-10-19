function helpPopupUndo() {
  const ui = DocumentApp.getUi();
  ui.alert(`You may have noticed that bFormat actions cannot be undone. The reason for this is that we are using the Docs API. The reason for using the Docs API is that there are certain operations that are only possible with the Doc API. 
However, it does have the disadvantage that you cannot undo those functions. Use bFormat functions with care. If necessary you can revert to an earlier version of the document using the File > Version History.`);
}

function onOpen(e) {
  let tryToRetrieveProperties, addonMenuFunction;

  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    tryToRetrieveProperties = false;
    addonMenuFunction = '.runUpdateMenu';
  } else {
    tryToRetrieveProperties = true;
    addonMenuFunction = '.run';
  }

  const thisDocStyle = getThisDocumentStyle(tryToRetrieveProperties);

  // Apply more styles submenu
  const subMenu = DocumentApp.getUi().createMenu('Apply more styles');
  let selectedStyleMarker = '';
  if (tryToRetrieveProperties === false) {
    subMenu.addItem('Activate add-on', 'activateAddOn');
  } else {
    for (let styleName in styles) {
      if (styleName == thisDocStyle.style) {
        selectedStyleMarker = thisDocStyle.marker + ' ';
      } else {
        selectedStyleMarker = '';
      }
      subMenu.addItem(selectedStyleMarker + styles[styleName]['name'], styleName);
    }
  }
  // End. Apply more styles submenu

  const menu = DocumentApp.getUi().createMenu('bFormat')
  for (let menuItemFunc in addOnMenu) {
    menu.addItem(addOnMenu[menuItemFunc]['txtMenuName'], 'addOnMenu' + '.' + menuItemFunc + addonMenuFunction);
    if (addOnMenu[menuItemFunc].separatorBelow === true) {
      menu.addSeparator();
    }
    if (addOnMenu[menuItemFunc].subMenuBelow === true) {
      menu.addItem('Apply style: ' + styles[thisDocStyle.domainBasedStyle]['name'], thisDocStyle.domainBasedStyle);
      menu.addSubMenu(subMenu);
    }
  }
  menu.addToUi();
}

function activateAddOn() {
  onOpen();
}

function runUpdateMenu(obj) {
  onOpen();
  obj.run.call();
}

const addOnMenu = {
  s0: {
    txtMenuName: 'Why can I not use undo?',
    run: helpPopupUndo,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s1: {
    txtMenuName: 'Help for setting default styles',
    run: setDefaultStylesManually,
    runUpdateMenu: function () { runUpdateMenu(this); },
    separatorBelow: true,
    subMenuBelow: true
  },
  s2: {
    txtMenuName: 'Format text like Heading 1',
    run: formatTextLikeH1,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s3: {
    txtMenuName: 'Reformat headings for tables, figures, boxes',
    run: reformatHeadings5and6,
    runUpdateMenu: function () { runUpdateMenu(this); },
    separatorBelow: true
  },
  s4: {
    txtMenuName: 'Insert box',
    run: insertBox,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s5: {
    txtMenuName: 'Format this box',
    run: formatBox,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s6: {
    txtMenuName: 'Add left-border to paragraph',
    run: leftBorderParagraph,
    runUpdateMenu: function () { runUpdateMenu(this); },
    separatorBelow: true
  },
  s7: {
    txtMenuName: 'Insert table 2x2',
    run: insertTable2x2,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s8: {
    txtMenuName: 'Insert table 3x3',
    run: insertTable3x3,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s9: {
    txtMenuName: 'Insert table 4x4',
    run: insertTable4x4,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s10: {
    txtMenuName: 'Format this table',
    run: formatTableNoBold,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s11: {
    txtMenuName: 'Format this table (all bold)',
    run: formatTable,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s12: {
    txtMenuName: 'Format this table (basic)',
    run: formatTableBasic,
    runUpdateMenu: function () { runUpdateMenu(this); },
    separatorBelow: true
  },
  s13: {
    txtMenuName: 'Insert figure/image (style 1)',
    run: insertFigure1,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s14: {
    txtMenuName: 'Insert figure/image (style 2)',
    run: insertFigure2,
    runUpdateMenu: function () { runUpdateMenu(this); },
    separatorBelow: true
  },
  s15: {
    txtMenuName: 'Insert pull quote',
    run: insertPullQuote,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s16: {
    txtMenuName: 'Insert extracted quote',
    run: insertExtractedQuote,
    runUpdateMenu: function () { runUpdateMenu(this); },
    separatorBelow: true
  },

  s17: {
    txtMenuName: 'Format lists',
    run: formatListsPart1,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s18: {
    txtMenuName: 'Remove underline from hyperlinks',
    run: removeUnderlineFromHyperlinks,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s19: {
    txtMenuName: 'Replace non-smart quotes with smart quotes',
    run: replaceNonSmartWithSmartQuotes,
    runUpdateMenu: function () { runUpdateMenu(this); },
    separatorBelow: true
  },
  s20: {
    txtMenuName: 'Format header',
    run: formatHeader,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s21: {
    txtMenuName: 'Update footer',
    run: formatFooter,
    runUpdateMenu: function () { runUpdateMenu(this); }
  },
  s22: {
    txtMenuName: 'Use default margins (Report)',
    run: defaultMargins,
    runUpdateMenu: function () { runUpdateMenu(this); }
  }
};