// OPS
/*
const CONFIG = {
  SPREADSHEET_ID: '1KFnjVbU4P0_YlUsKOIOhtIozduQwCSfr4v78J2o3GaM',
  SHEET_NAMES: {
    LEADS: 'leads',
    GENERATED_FILES_FOLDER_ID: '14lLfaWmCu949BO1VcrSlMW1JdS8yCvuA', // <-- ĮRAŠYKITE SAVO ARCHYVO/SUGENERUOTŲ FAILŲ APLANKO ID ČIA
    CONFIGURATION: 'configuration',
    PROPOSAL_CONFIG: 'config_pasiulymas',
    PRICES: 'Kainos',
    PROPOSAL_CALCULATION_PREFIX: 'pasiulymas',
    PROPOSAL_TEMPLATE_PREFIX: 'template_pasiulymas',
    NEW_PROPOSAL_SHEET: 'Pasiūlymas'
  }
}; */

// TEST
const CONFIG = {
  SPREADSHEET_ID: '1QXWE2WgukqOFWZBwL1aYfS-C9dDzrdzyz9E6iBlV24o',
  SHEET_NAMES: {
    LEADS: 'leads',
    GENERATED_FILES_FOLDER_ID: '1kz8ZFwQ61AemThG72rAPyoTl6dRoDRx7', // <-- ĮRAŠYKITE SAVO ARCHYVO/SUGENERUOTŲ FAILŲ APLANKO ID ČIA
    CONFIGURATION: 'configuration',
    PROPOSAL_CONFIG: 'config_pasiulymas',
    PRICES: 'Kainos',
    PROPOSAL_CALCULATION_PREFIX: 'pasiulymas',
    PROPOSAL_TEMPLATE_PREFIX: 'template_pasiulymas',
    NEW_PROPOSAL_SHEET: 'Pasiūlymas'
  },
  PROPOSAL_AUTO_DISPLAY_CONFIG: 'config_pasiulymas_autoatvaizdavimas'
};

function doGet() {
  try {
    var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet1 = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.LEADS);
    if (!sheet1) {
      return HtmlService.createHtmlOutput('<p>Klaida: Trūksta "leads" lapo.</p>');
    }

    var config = loadConfiguration(spreadsheet);
    var { selectedColumns, mappedNames, editableColumns, optionsColumns, dateButtonColumns, rowCounts, dateTimeColumns, dateColumns, datePickerColumns, columnPositions, pasiulymoReiksmes, formulas, proposal } = config;
  
  var data = sheet1.getDataRange().getValues();
  var headers = data[0].map(header => header.toString().trim().toLowerCase()); // Trim and normalize headers
  
  // Validate selectedColumns against headers
  var invalidColumns = selectedColumns.filter(col => !headers.includes(col.toLowerCase()));
  if (invalidColumns.length > 0) {
    Logger.log('Invalid columns in stulpeliai_rodyti: ' + invalidColumns.join(', '));
    return HtmlService.createHtmlOutput('<p>Klaida: Netinkami stulpeliai configuration lape: ' + invalidColumns.join(', ') + '. Patikrinkite "stulpeliai_rodyti" stulpelį configuration lape ir įsitikinkite, kad visi išvardyti stulpeliai egzistuoja leads lapo antraštėse.</p>');
  }
  
  // Validate proposalSelectedColumns against headers
  // Check if at least the first numbered version of each column exists (e.g., "pasirinkite Kw1")
  var invalidProposalColumns = proposal.selectedColumns.filter(col => !headers.includes((col + '1').toLowerCase()));
  if (invalidProposalColumns.length > 0) {
    Logger.log('Invalid columns in config_pasiulymas: ' + invalidProposalColumns.join(', '));
    return HtmlService.createHtmlOutput('<p>Klaida: Netinkami stulpeliai config_pasiulymas lape: ' + invalidProposalColumns.join(', ') + 
                                        '. Patikrinkite "pasiulymas_stulpeliai_rodyti" stulpelį config_pasiulymas lape ir įsitikinkite, kad atitinkami sunumeruoti stulpeliai (pvz., "' + 
                                        invalidProposalColumns[0] + '1") egzistuoja leads lapo antraštėse.</p>');
  }
  
  function getCellValueOrFormula(rangeNotation, rowIndex) {
    if (!rangeNotation) return [];
    rangeNotation = rangeNotation.toString().replace(/^=/, '').trim(); // Strip leading '=' and trim
    try {
      var range = spreadsheet.getRange(rangeNotation);
      var dataValidation = range.getDataValidation();
      if (dataValidation && dataValidation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
        var validationRange = dataValidation.getCriteriaValues()[0];
        var values = validationRange.getValues().flat().filter(val => val !== '').map(val => val.toString().trim());
        return values;
      } else if (dataValidation && dataValidation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        var values = dataValidation.getCriteriaValues()[0];
        return values.split(',').map(val => val.trim()).filter(val => val !== '');
      }
      if (rangeNotation.includes('!')) {
        var parts = rangeNotation.split('!');
        var sheetName = parts[0].trim();
        var cellRange = parts[1].trim();
        return spreadsheet.getSheetByName(sheetName).getRange(cellRange).getValues().flat().filter(val => val !== '').map(val => val.toString().trim());
      }
      var value = range.getValue();
      return value ? [value.toString().trim()] : [];
    } catch (e) {
      Logger.log('Error retrieving cell value or data validation for ' + rangeNotation + ': ' + e);
      return [];
    }
  }
  
  var kainosValues = {
    'A21': spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.PRICES).getRange('A21').getValue().toString().trim(),
    'A23': spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.PRICES).getRange('A23').getValue().toString().trim(),
    'A24': spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.PRICES).getRange('A24').getValue().toString().trim(),
    'A38': spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.PRICES).getRange('A38').getValue().toString().trim(),
    'A42': spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.PRICES).getRange('A42').getValue().toString().trim()
  };
  
  var columnIndices = [];
  var displayNames = [];
  var isEditable = [];
  var columnOptions = [];
  var hasDateButton = [];
  var rowCountsConfig = [];
  var isDateTimeColumn = [];
  var isDateColumn = [];
  var hasDatePicker = [];
  var positions = [];
  var pasiulymoOptions = [];
  var formulaOptions = [];
  var columnNames = [];
  selectedColumns.forEach(function(colName, index) {
    var colIndex = headers.indexOf(colName.toLowerCase());
    if (colIndex !== -1) {
      columnIndices.push(colIndex);
      displayNames.push(mappedNames[index] ? mappedNames[index].trim() : colName);
      isEditable.push(editableColumns[index] === 'x' || dateButtonColumns[index] === 'x' || datePickerColumns[index] === 'x');
      var cleanedOptions = optionsColumns[index] ? optionsColumns[index].split(',').map(opt => opt.trim()).filter(opt => opt !== '') : [];
      columnOptions.push(cleanedOptions);
      hasDateButton.push(dateButtonColumns[index] === 'x');
      rowCountsConfig.push(rowCounts[index] && !isNaN(rowCounts[index]) ? parseInt(rowCounts[index]) : 1);
      isDateTimeColumn.push(dateTimeColumns[index] === 'x');
      isDateColumn.push(dateColumns[index] === 'x');
      hasDatePicker.push(datePickerColumns[index] === 'x');
      positions.push(columnPositions[index] || 1);
      
      // Nustatome parinktis. `issokusi_reiksmes` turi aukščiausią prioritetą.
      let opts = cleanedOptions.length > 0 ? cleanedOptions : 
                 (pasiulymoReiksmes[index] ? getCellValueOrFormula(pasiulymoReiksmes[index], 0) : 
                 (formulas[index] ? getCellValueOrFormula(formulas[index], 0) : []));
      pasiulymoOptions.push(opts); // Naudosime šį masyvą visoms parinktims

      columnNames.push(colName.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase());
    }
  });
  
  if (columnIndices.length === 0) {
    return HtmlService.createHtmlOutput('<p>Klaida: Nerasta tinkamų stulpelių leads antraštėse. Patikrinkite "stulpeliai_rodyti" stulpelį configuration lape.</p>');
  }
  
  var invalidPositions = positions.filter(pos => pos !== 1 && pos !== 2);
  if (invalidPositions.length > 0) {
    Logger.log('Invalid positions in Stulpeliai: ' + invalidPositions.join(', '));
  }
  
  var group1Columns = [];
  var group2Columns = [];
  columnIndices.forEach(function(idx, i) {
    if (positions[i] === 1) {
      group1Columns.push({ index: idx, displayName: displayNames[i], editable: isEditable[i], options: pasiulymoOptions[i], dateButton: hasDateButton[i], rowCount: rowCountsConfig[i], isDateTime: isDateTimeColumn[i], isDate: isDateColumn[i], hasDatePicker: hasDatePicker[i], columnName: columnNames[i] });
    } else if (positions[i] === 2) {
      group2Columns.push({ index: idx, displayName: displayNames[i], editable: isEditable[i], options: pasiulymoOptions[i], dateButton: hasDateButton[i], rowCount: rowCountsConfig[i], isDateTime: isDateTimeColumn[i], isDate: isDateColumn[i], hasDatePicker: hasDatePicker[i], columnName: columnNames[i] });
    }
  });
  
  // Process columns for proposal modal
  var proposalColumnIndices = [];
  var proposalDisplayNames = [];
  var proposalIsEditable = [];
  var proposalColumnOptions = [];
  var proposalColumnNames = [];
  var allProposalColumns = [];

  // Ciklas per visus galimus stulpelius, kad rastume sunumeruotus variantus
  headers.forEach(function(header, colIndex) {
    var match = header.match(/^([a-z\s_]+)(\d)$/i); // Pvz., "pasirinkite kw1" -> ["pasirinkite kw1", "pasirinkite kw", "1"]
    if (!match) return;

    var baseName = match[1].trim(); // "pasirinkite kw"
    var number = match[2]; // "1"
    var configIndex = proposal.selectedColumns.findIndex(c => c.toLowerCase() === baseName.toLowerCase());

    if (configIndex !== -1) {
      var dropdownLinkTemplate = proposal.dropdownLinkColumns[configIndex];
      var autoUpdateCellTemplate = proposal.autoUpdateColumns[configIndex];
      var dropdownOptions = [];

      if (dropdownLinkTemplate) {
        var dynamicRange = dropdownLinkTemplate.replace('pasiulymas!', CONFIG.SHEET_NAMES.PROPOSAL_CALCULATION_PREFIX + number + '!');
        dropdownOptions = getCellValueOrFormula(dynamicRange, 0);
      }

      allProposalColumns.push({
        index: colIndex,
        displayName: baseName,
        editable: proposal.editableColumns[configIndex] === 'x',
        options: dropdownOptions,
        columnName: header.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase(), // pvz., pasirinkite_kw1
        // Priskiriame auto-update reikšmę, jei ji yra nurodyta konfigūracijoje.
        autoUpdateCell: autoUpdateCellTemplate || ''
      });
    }
  });
  
  if (allProposalColumns.length === 0) {
    return HtmlService.createHtmlOutput('<p>Klaida: Nerasta tinkamų stulpelių leads antraštėse pagal config_pasiulymas. Patikrinkite "stulpeliai_rodyti" stulpelį config_pasiulymas lape.</p>');
  }

  // Base structure for the modal (uses base names)
  var proposalColumns = proposal.selectedColumns.map((colName, i) => ({
    displayName: colName,
    editable: proposal.editableColumns[i] === 'x',
    options: [], // This is not used, options are taken from allProposalColumns
    dateButton: false,
    rowCount: 1,
    isDateTime: false,
    isDate: false,
    hasDatePicker: false,
    autoUpdateCell: proposal.autoUpdateColumns[i] || '', // Pridedame auto-update konfigūraciją
    // autoUpdateCell is no longer needed here, it's in allProposalColumns
    formulaOptions: [],
    columnName: colName.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase()
  }));
  
  // Get pasiulymu_kiekis options and index
  var pasiulymuKiekisIndex = headers.indexOf('pasiulymu_kiekis'); // This is a column name, not a sheet name
  var pasiulymuKiekisOptions = ['1', '2', '3']; // Leidžiame iki 3 pasiūlymų
  
  Logger.log('group1Columns: ' + JSON.stringify(group1Columns.map(col => ({ columnName: col.columnName, displayName: col.displayName }))));
  Logger.log('group2Columns: ' + JSON.stringify(group2Columns.map(col => ({ columnName: col.columnName, displayName: col.displayName }))));
  Logger.log('allProposalColumns: ' + JSON.stringify(allProposalColumns.map(col => ({ columnName: col.columnName, index: col.index }))));
  function formatToVilniusDateTime(cell) {
    if (!cell) return '';
    try {
      var date;
      if (cell instanceof Date) {
        date = cell;
      } else if (typeof cell === 'string') {
        if (cell.match(/^\d{4}-\d{2}-\d{2}$/)) {
          date = new Date(cell + 'T00:00:00');
        } else if (cell.match(/^\d{4}-\d{2}-\d{2} \d{1,2}:\d{2}$/)) {
          date = new Date(cell.replace(' ', 'T') + ':00');
        } else if (cell.includes('T')) {
          date = new Date(cell);
        } else {
          return cell;
        }
      } else {
        return cell;
      }
      return Utilities.formatDate(date, 'Europe/Vilnius', 'yyyy-MM-dd HH:mm');
    } catch (e) {
      Logger.log('Error formatting date-time: ' + cell + ', Error: ' + e);
      return cell;
    }
  }
  
  function formatToVilniusDate(cell) {
    if (!cell) return '';
    try {
      var date;
      if (cell instanceof Date) {
        date = cell;
      } else if (typeof cell === 'string') {
        if (cell.match(/^\d{4}-\d{2}-\d{2}$/)) {
          date = new Date(cell + 'T00:00:00');
        } else if (cell.match(/^\d{4}-\d{2}-\d{2} \d{1,2}:\d{2}$/)) {
          date = new Date(cell.replace(' ', 'T') + ':00');
        } else if (cell.includes('T')) {
          date = new Date(cell);
        } else {
          return cell;
        }
      } else {
        return cell;
      }
      return Utilities.formatDate(date, 'Europe/Vilnius', 'yyyy-MM-dd');
    } catch (e) {
      Logger.log('Error formatting date: ' + cell + ', Error: ' + e);
      return cell;
    }
  }
  
  function toISO8601(cell) {
    if (!cell) return '';
    try {
      var date;
      if (cell.match(/^\d{4}-\d{2}-\d{2}$/)) {
        date = new Date(cell + 'T00:00:00');
      } else if (cell.match(/^\d{4}-\d{2}-\d{2} \d{1,2}:\d{2}$/)) {
        date = new Date(cell.replace(' ', 'T') + ':00');
      } else if (cell.includes('T')) {
        date = new Date(cell);
      } else {
        return cell;
      }
      return Utilities.formatDate(date, 'Europe/Vilnius', "yyyy-MM-dd'T'HH:mm:ssZ").replace(/(\d{2})(\d{2})$/, '$1:$2');
    } catch (e) {
      Logger.log('Error converting to ISO 8601: ' + cell + ', Error: ' + e);
      return cell;
    }
  }
  
  var formattedData = data.map(function(row, rowIndex) {
    if (rowIndex === 0) return row;
    return row.map(function(cell, colIndex) {
      var colName = headers[colIndex];
      var colIndexInSelected = selectedColumns.indexOf(colName);
      var isDateTime = colIndexInSelected !== -1 ? dateTimeColumns[colIndexInSelected] === 'x' : false;
      var isDate = colIndexInSelected !== -1 ? dateColumns[colIndexInSelected] === 'x' : false;
      if (typeof cell === 'string') {
        cell = cell.replace(/\\:/g, ':');
      }
      if (colName === 'phone_number' && typeof cell === 'string') {
        cell = cell.replace(/^p:/, '');
      }
      if (isDateTime) {
        return formatToVilniusDateTime(cell);
      } else if (isDate) {
        return formatToVilniusDate(cell);
      }
      return cell || '';
    });
  });
  var html = HtmlService.createTemplateFromFile('View');
  html.formattedData = formattedData;
  html.displayNames = displayNames;
  html.columnIndices = columnIndices;
  html.group1Columns = group1Columns;
  html.group2Columns = group2Columns;
  html.proposalColumns = proposalColumns;
  html.allProposalColumns = allProposalColumns; // Pass all numbered columns to the view
  html.autoDisplayConfig = proposal.autoDisplay; // Perduodame auto-atvaizdavimo konfigūraciją
  html.pasiulymuKiekisOptions = pasiulymuKiekisOptions;
  html.pasiulymuKiekisIndex = pasiulymuKiekisIndex;
  html.kainosValues = kainosValues;
  // Surandame ir perduodame 'id' stulpelio indeksą
  var idIndex = headers.indexOf('id');
  html.idIndex = idIndex;

  var fullNameIndex = headers.indexOf('full_name');
  html.fullNameIndex = fullNameIndex;

  var emailIndex = headers.indexOf('email');
  html.emailIndex = emailIndex;

  var linksIndex = headers.indexOf('pasiulymu_nuorodos');
  html.linksIndex = linksIndex;

  return html.evaluate().setTitle('Facebook Leads Duomenys').setWidth(1000).setHeight(600);
  } catch (e) {
    Logger.log('Klaida vykdant doGet: ' + e.toString());
    return HtmlService.createHtmlOutput('<p>Įvyko kritinė klaida: ' + e.toString() + '</p>');
  }
}

function updateSheet1(rowIndex, updates) {
  var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.LEADS);
  if (!sheet) {
    throw new Error('Lapas "' + CONFIG.SHEET_NAMES.LEADS + '" nerastas.');
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  for (var colIndex in updates) {
    var update = updates[colIndex];
    var value = update.value;
    if (update.isDate && value) {
      try {
        var date;
        if (value.match(/^\d{4}-\d{2}-\d{2}$/)) {
          date = new Date(value + 'T00:00:00');
        } else if (value.match(/^\d{4}-\d{2}-\d{2} \d{1,2}:\d{2}$/)) {
          date = new Date(value.replace(' ', 'T') + ':00');
        } else if (value.includes('T')) {
          date = new Date(value);
        } else {
          value = value;
        }
        if (date && !isNaN(date.getTime())) {
          value = Utilities.formatDate(date, 'Europe/Vilnius', "yyyy-MM-dd'T'HH:mm:ssZ").replace(/(\d{2})(\d{2})$/, '$1:$2');
        }
      } catch (e) {
        Logger.log('Error converting to ISO 8601 in updateSheet1: ' + value + ', Error: ' + e);
      }
    }
    sheet.getRange(rowIndex + 1, parseInt(colIndex) + 1).setValue(value);
  }
}

/**
 * Updates a specific cell in one of the proposal calculation sheets.
 * This function is called from the client-side when an input with auto-update configured is changed.
 *
 * @param {number} sheetNumber The number of the proposal (1, 2, or 3).
 * @param {string} cellAddress The A1 notation of the cell to update (e.g., 'C7').
 * @param {any} value The new value to set in the cell.
 */
function updateProposalSheetCell(sheetNumber, cellAddress, value) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheetName = CONFIG.SHEET_NAMES.PROPOSAL_CALCULATION_PREFIX + sheetNumber;
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('Skaičiavimo lapas "' + sheetName + '" nerastas.');
    }

    sheet.getRange(cellAddress).setValue(value);
    SpreadsheetApp.flush(); // Svarbu: laukiame, kol formulės persiskaičiuos

    // Po atnaujinimo, nuskaitome ir grąžiname auto-atvaizdavimo reikšmes
    const config = loadConfiguration(spreadsheet);
    const autoDisplayConfig = config.proposal.autoDisplay;

    if (!autoDisplayConfig || autoDisplayConfig.length === 0) {
      return { success: true, autoDisplayValues: [] };
    }

    const newValues = autoDisplayConfig.map(item => {
      const pavadinimas = sheet.getRange(item.pavadinimasRef.split('!')[1]).getDisplayValue();
      const reiksme = sheet.getRange(item.reiksmeRef.split('!')[1]).getDisplayValue();
      return { pavadinimas, reiksme };
    });

    return { success: true, autoDisplayValues: newValues };
  } catch (e) {
    Logger.log('Klaida vykdant updateProposalSheetCell: ' + e.toString());
    throw new Error('Nepavyko atnaujinti langelio: ' + e.message);
  }
}

function syncAndGetInitialData(sheetNumber, dataToSync) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheetName = CONFIG.SHEET_NAMES.PROPOSAL_CALCULATION_PREFIX + sheetNumber;
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('Skaičiavimo lapas "' + sheetName + '" nerastas.');
    }

    // 1. Atnaujiname langelius su pateiktais duomenimis
    for (const item of dataToSync) {
      sheet.getRange(item.cellAddress).setValue(item.value);
    }

    SpreadsheetApp.flush(); // Svarbu: laukiame, kol formulės persiskaičiuos

    // 2. Nuskaitome ir grąžiname auto-atvaizdavimo reikšmes
    const config = loadConfiguration(spreadsheet);
    const autoDisplayConfig = config.proposal.autoDisplay;
    if (!autoDisplayConfig || autoDisplayConfig.length === 0) {
      return { success: true, autoDisplayValues: [] };
    }

    const newValues = autoDisplayConfig.map(item => {
      const pavadinimas = sheet.getRange(item.pavadinimasRef.split('!')[1]).getDisplayValue();
      const reiksme = sheet.getRange(item.reiksmeRef.split('!')[1]).getDisplayValue();
      return { pavadinimas, reiksme };
    });

    return { success: true, autoDisplayValues: newValues };
  } catch (e) {
    Logger.log('Klaida vykdant syncAndGetInitialData: ' + e.toString());
    throw new Error('Nepavyko sinchronizuoti duomenų: ' + e.message);
  }
}