function helpPopupUndo() {
  const ui = DocumentApp.getUi();
  ui.alert(`You may have noticed that bFormat actions cannot be undone. The reason for this is that we are using the Docs API. The reason for using the Docs API is that there are certain operations that are only possible with the Doc API. 
However, it does have the disadvantage that you cannot undo those functions. Use bFormat functions with care. If necessary you can revert to an earlier version of the document using the File > Version History.`);
}

function onOpen(e) {
  universal_bFormat_menu(e, 'menu').addToUi();
  /*  let tryToRetrieveProperties, addonMenuFunction;
  
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
    */
}

function universal_bFormat_menu(e, returnType = 'menu') {
  let tryToRetrieveProperties, methodName;

  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    tryToRetrieveProperties = false;
    methodName = 'runUpdateMenu';
  } else {
    tryToRetrieveProperties = true;
    // methodName = 'run';
    methodName = 'runUpdateMenu';
  }

  const thisDocStyle = getThisDocumentStyle(tryToRetrieveProperties);

  const menuStructure = {
    title: 'bFormat',
    items: []
  };

  // Build "Apply more styles" submenu
  const applyMoreStylesSubmenu = {
    type: 'submenu',
    label: 'Apply more styles',
    items: []
  };

  let selectedStyleMarker = '';
  if (tryToRetrieveProperties === false) {
    applyMoreStylesSubmenu.items.push({
      type: 'item',
      label: 'Activate add-on',
      functionName: 'activateAddOn'
    });
  } else {
    for (let styleName in styles) {
      if (styleName == thisDocStyle.style) {
        selectedStyleMarker = thisDocStyle.marker + ' ';
      } else {
        selectedStyleMarker = '';
      }
      applyMoreStylesSubmenu.items.push({
        type: 'item',
        label: selectedStyleMarker + styles[styleName]['name'],
        functionName: styleName
      });
    }
  }

  // Build main menu items
  for (let menuItemFunc in addOnMenu) {
    menuStructure.items.push({
      type: 'item',
      label: addOnMenu[menuItemFunc]['txtMenuName'],
      functionName: 'addOnMenu',
      functionParams: [menuItemFunc, methodName]
    });

    if (addOnMenu[menuItemFunc].separatorBelow === true) {
      menuStructure.items.push({ type: 'separator' });
    }

    if (addOnMenu[menuItemFunc].subMenuBelow === true) {
      menuStructure.items.push({
        type: 'item',
        label: 'Apply style: ' + styles[thisDocStyle.domainBasedStyle]['name'],
        functionName: thisDocStyle.domainBasedStyle
      });
      menuStructure.items.push(applyMoreStylesSubmenu);
    }
  }

  if (returnType === 'data') {
    return menuStructure;
  } else {
    return buildUIMenu(menuStructure);
  }
}

function buildUIMenu(menuStructure) {
  menuStructure.items.unshift({
    "type": "item",
    "label": "Open menu in sidebar 🚀",
    "functionName": "showSidebarMenu"
  },
    { type: 'separator' });

  const menu = DocumentApp.getUi().createMenu(menuStructure.title);
  
  menuStructure.items.forEach(item => {
    if (item.type === 'separator') {
      menu.addSeparator();
    } else if (item.type === 'item') {
      // Handle items with parameters (nested object methods)
      if (item.functionParams && item.functionParams.length > 0) {
        if (item.functionParams.length === 2) {
          // Full reconstruction: objectName.property.method
          menu.addItem(item.label, item.functionName + '.' + item.functionParams[0] + '.' + item.functionParams[1]);
        } else if (item.functionParams.length === 1) {
          // Single parameter: objectName.property.run (default)
          menu.addItem(item.label, item.functionName + '.' + item.functionParams[0] + '.run');
        }
      } else {
        menu.addItem(item.label, item.functionName);
      }
    } else if (item.type === 'submenu') {
      const submenu = DocumentApp.getUi().createMenu(item.label);
      item.items.forEach(subItem => {
        if (subItem.type === 'separator') {
          submenu.addSeparator();
        } else if (subItem.type === 'item') {
          // Handle subitems with parameters
          if (subItem.functionParams && subItem.functionParams.length > 0) {
            if (subItem.functionParams.length === 2) {
              submenu.addItem(subItem.label, subItem.functionName + '.' + subItem.functionParams[0] + '.' + subItem.functionParams[1]);
            } else if (subItem.functionParams.length === 1) {
              submenu.addItem(subItem.label, subItem.functionName + '.' + subItem.functionParams[0] + '.run');
            }
          } else {
            submenu.addItem(subItem.label, subItem.functionName);
          }
        }
      });
      menu.addSubMenu(submenu);
    }
  });
  
  return menu;
}

// Dispatcher function for sidebar calls to addOnMenu object
function addOnMenu_sidebar(menuItemFunc, methodName = 'run') {
  if (addOnMenu[menuItemFunc] && addOnMenu[menuItemFunc][methodName]) {
    addOnMenu[menuItemFunc][methodName]();
  } else {
    throw new Error('Invalid menu item or method: ' + menuItemFunc + '.' + methodName);
  }
}

function showSidebarMenu() {
  const template = HtmlService.createTemplateFromFile('000 Menu sidebar');
  const menuStructure = universal_bFormat_menu(null, 'data');
  template.menuStructureJson = JSON.stringify(menuStructure);

  const html = template.evaluate()
    .setTitle(menuStructure.title + ' menu');

  DocumentApp.getUi().showSidebar(html);
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
  },
  s23: {
    txtMenuName: 'Clear internal broken link markers',
    run: clearInternalLinkMarkers,
    runUpdateMenu: function () { runUpdateMenu(this); }
  }
};