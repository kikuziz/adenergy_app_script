// OPS
/*
const CONFIG = {
  SPREADSHEET_ID: '1KFnjVbU4P0_YlUsKOIOhtIozduQwCSfr4v78J2o3GaM',
  TITLE: 'OPS v1.00',
  SHEET_NAMES: {
    LEADS: 'leads',
    GENERATED_FILES_FOLDER_ID: '1zLCtWiPzcuX1m6RTt6q4zibSdz6BJjKm', // <-- ĮRAŠYKITE SAVO ARCHYVO/SUGENERUOTŲ FAILŲ APLANKO ID ČIA
    CONFIGURATION: 'configuration',
    PROPOSAL_CONFIG: 'config_pasiulymas',
    PRICES: 'Kainos',
    PROPOSAL_CALCULATION_PREFIX: 'pasiulymas',
    PROPOSAL_TEMPLATE_PREFIX: 'template_pasiulymas',
    NEW_PROPOSAL_SHEET: 'Pasiūlymas',
    PROPOSAL_AUTO_DISPLAY_CONFIG: 'config_pasiulymas_autoatvaizdavimas',
    EMAIL_CONFIG: 'config_mail',
    NEW_PROPOSAL_SHEET: 'Pasiūlymas',
    PROPOSAL_AUTO_DISPLAY_CONFIG: 'config_pasiulymas_autoatvaizdavimas',
    EMAIL_CONFIG: 'config_mail'
  }
  
};  */

// TEST
const CONFIG = {
  SPREADSHEET_ID: '1QXWE2WgukqOFWZBwL1aYfS-C9dDzrdzyz9E6iBlV24o',
  TITLE: 'TEST v1.01',
  SHEET_NAMES: {
    LEADS: 'leads',
    GENERATED_FILES_FOLDER_ID: '1kz8ZFwQ61AemThG72rAPyoTl6dRoDRx7', // <-- ĮRAŠYKITE SAVO ARCHYVO/SUGENERUOTŲ FAILŲ APLANKO ID ČIA
    CONFIGURATION: 'configuration',
    PROPOSAL_CONFIG: 'config_pasiulymas',
    PRICES: 'Kainos',
    PROPOSAL_CALCULATION_PREFIX: 'pasiulymas',
    PROPOSAL_TEMPLATE_PREFIX: 'template_pasiulymas',
    NEW_PROPOSAL_SHEET: 'Pasiūlymas',
    PROPOSAL_AUTO_DISPLAY_CONFIG: 'config_pasiulymas_autoatvaizdavimas',
    EMAIL_CONFIG: 'config_mail'
  }
};

function doGet() {
  try {
    var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet1 = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.LEADS);
    if (!sheet1) {
      return HtmlService.createHtmlOutput('<p>Klaida: Trūksta "leads" lapo.</p>');
    }

    var config = loadConfiguration(spreadsheet);
    var { selectedColumns, mappedNames, editableColumns, optionsColumns, dateButtonColumns, rowCounts, dateTimeColumns, dateColumns, datePickerColumns, columnPositions, pasiulymoReiksmes, proposal } = config;
  
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
  var originalColumnNames = []; // Pridedame originalių pavadinimų masyvą
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
                 (pasiulymoReiksmes[index] ? getCellValueOrFormula(pasiulymoReiksmes[index], 0) : []);
      pasiulymoOptions.push(opts); // Naudosime šį masyvą visoms parinktims

      columnNames.push(colName.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase());
      originalColumnNames.push(colName); // Išsaugome originalų pavadinimą
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
      group1Columns.push({ index: idx, displayName: displayNames[i], editable: isEditable[i], options: pasiulymoOptions[i], dateButton: hasDateButton[i], rowCount: rowCountsConfig[i], isDateTime: isDateTimeColumn[i], isDate: isDateColumn[i], hasDatePicker: hasDatePicker[i], columnName: columnNames[i], headerName: originalColumnNames[i] });
    } else if (positions[i] === 2) {
      group2Columns.push({ index: idx, displayName: displayNames[i], editable: isEditable[i], options: pasiulymoOptions[i], dateButton: hasDateButton[i], rowCount: rowCountsConfig[i], isDateTime: isDateTimeColumn[i], isDate: isDateColumn[i], hasDatePicker: hasDatePicker[i], columnName: columnNames[i], headerName: originalColumnNames[i] });
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

  return html.evaluate().setTitle(CONFIG.TITLE).setWidth(1000).setHeight(600);
  } catch (e) {
    Logger.log('Klaida vykdant doGet: ' + e.toString());
    return HtmlService.createHtmlOutput('<p>Įvyko kritinė klaida: ' + e.toString() + '</p>');
  }
}

function updateleadSheet(rowIndex, updates) {
  Logger.log('updateleadSheet called with rowIndex: ' + rowIndex + ', updates: ' + JSON.stringify(updates));
  var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var leadSheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.LEADS);
  if (!leadSheet) {
    throw new Error('Lapas "' + CONFIG.SHEET_NAMES.LEADS + '" nerastas.');
  }
  var headers = leadSheet.getRange(1, 1, 1, leadSheet.getLastColumn()).getValues()[0];
  
  // Sukuriame antraščių žemėlapį (header name -> index) greitesnei paieškai
  var headerMap = {};
  headers.forEach((header, i) => {
    headerMap[header.toLowerCase()] = i;
  });

  for (var headerName in updates) {
    var update = updates[headerName];
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
        Logger.log('Error converting to ISO 8601 in updateleadSheet: ' + value + ', Error: ' + e);
      }
    }
    
    var colIndex = headerMap[headerName.toLowerCase()];
    if (colIndex !== undefined) {
      leadSheet.getRange(rowIndex + 1, colIndex + 1).setValue(value);
    } else {
      Logger.log('WARNING: Column "' + headerName + '" not found in leadSheet. Skipping update.');
    }
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
  Logger.log('updateProposalSheetCell called with sheetNumber: ' + sheetNumber + ', cellAddress: ' + cellAddress + ', value: ' + value);
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheetName = CONFIG.SHEET_NAMES.PROPOSAL_CALCULATION_PREFIX + sheetNumber;
    const calculationSheet = spreadsheet.getSheetByName(sheetName);
    
    if (!calculationSheet) {
      throw new Error('Skaičiavimo lapas "' + sheetName + '" nerastas.');
    }

    calculationSheet.getRange(cellAddress).setValue(value);
    SpreadsheetApp.flush(); // Svarbu: laukiame, kol formulės persiskaičiuos

    // Po atnaujinimo, nuskaitome ir grąžiname auto-atvaizdavimo reikšmes
    const config = loadConfiguration(spreadsheet);
    const autoDisplayConfig = config.proposal.autoDisplay;

    if (!autoDisplayConfig || autoDisplayConfig.length === 0) {
      return { success: true, autoDisplayValues: [] };
    }

    const newValues = autoDisplayConfig.map(item => {
      const pavadinimas = calculationSheet.getRange(item.pavadinimasRef.split('!')[1]).getDisplayValue();
      const reiksme = calculationSheet.getRange(item.reiksmeRef.split('!')[1]).getDisplayValue();
      return { pavadinimas, reiksme };
    });

    return { success: true, autoDisplayValues: newValues };
  } catch (e) {
    Logger.log('Klaida vykdant updateProposalSheetCell: ' + e.toString());
    throw new Error('Nepavyko atnaujinti langelio: ' + e.message);
  }
}

function syncAndGetInitialData(calculationSheetNumber, dataToSync) {
  Logger.log('syncAndGetInitialData called with calculationSheetNumber: ' + calculationSheetNumber+ ' and dataToSync: ' + JSON.stringify(dataToSync));
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheetName = CONFIG.SHEET_NAMES.PROPOSAL_CALCULATION_PREFIX + calculationSheetNumber;
    const calculationSheet = spreadsheet.getSheetByName(sheetName);
    if (!calculationSheet) {
      throw new Error('Skaičiavimo lapas "' + sheetName + '" nerastas.');
    }

    // 1. Atnaujiname langelius su pateiktais duomenimis
    for (const item of dataToSync) {
      calculationSheet.getRange(item.cellAddress).setValue(item.value);
    }

    SpreadsheetApp.flush(); // Svarbu: laukiame, kol formulės persiskaičiuos

    // 2. Nuskaitome ir grąžiname auto-atvaizdavimo reikšmes
    const config = loadConfiguration(spreadsheet);
    const autoDisplayConfig = config.proposal.autoDisplay;
    if (!autoDisplayConfig || autoDisplayConfig.length === 0) {
      return { success: true, autoDisplayValues: [] };
    }

    const newValues = autoDisplayConfig.map(item => {
      const pavadinimas = calculationSheet.getRange(item.pavadinimasRef.split('!')[1]).getDisplayValue();
      const reiksme = calculationSheet.getRange(item.reiksmeRef.split('!')[1]).getDisplayValue();
      return { pavadinimas, reiksme };
    });

    return { success: true, autoDisplayValues: newValues };
  } catch (e) {
    Logger.log('Klaida vykdant syncAndGetInitialData: ' + e.toString());
    throw new Error('Nepavyko sinchronizuoti duomenų: ' + e.message);
  }
}

function generateProposalDocument(uniqueId) {
    try {
        var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.LEADS);
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var allData = sheet.getDataRange().getValues();

        // Find the index of the column to be used as a unique ID (e.g., 'id')
        var uniqueIdColIndex = headers.map(h => h.toLowerCase()).indexOf('id');
        if (uniqueIdColIndex === -1) {
            throw new Error("Could not find the 'id' column, which is used as a unique identifier.");
        }

        // Find the row corresponding to the unique ID
        var rowIndex = -1;
        var data;
        allData.forEach((row, i) => {
            if (i > 0 && row[uniqueIdColIndex].toString() === uniqueId) {
                rowIndex = i; // i is the index in the allData array (starts from 0)
                data = row;
            }
        });

        if (!data) {
            throw new Error('Could not find row with ID: ' + uniqueId);
        }

        var rowData = {};
        headers.forEach((header, i) => {
            rowData[header] = data[i];
        });

        // A more reliable way to find the 'pasiulymu_kiekis' value, regardless of case
        var proposalCountKey = Object.keys(rowData).find(key => key.toLowerCase() === 'pasiulymu_kiekis');
        var proposalCount = 1; // Default if not found or value is invalid
        if (proposalCountKey && rowData[proposalCountKey]) {
            var countValue = parseInt(rowData[proposalCountKey], 10);
            if (!isNaN(countValue) && countValue > 0) {
                proposalCount = countValue;
            }
        }
        var templateName = CONFIG.SHEET_NAMES.PROPOSAL_TEMPLATE_PREFIX + proposalCount;

        Logger.log('Generating proposal for row: ' + (rowIndex + 1) + ' using template: ' + templateName);
        
        // 1. Create a new Google Sheets file with a name
        var newSpreadsheet = SpreadsheetApp.create('Ad Energy pasiūlymas ' + (rowData['full_name'] || uniqueId));
        
        // Move the generated file to the specified archive folder
        if (CONFIG.SHEET_NAMES.GENERATED_FILES_FOLDER_ID && CONFIG.SHEET_NAMES.GENERATED_FILES_FOLDER_ID !== 'JUSU_ARCHYVO_APLANKO_ID') {
            try {
                var file = DriveApp.getFileById(newSpreadsheet.getId());
                var targetFolder = DriveApp.getFolderById(CONFIG.SHEET_NAMES.GENERATED_FILES_FOLDER_ID);
                file.moveTo(targetFolder);
                Logger.log('Generated proposal file moved to folder: ' + targetFolder.getName());
            } catch (e) {
                Logger.log('WARNING: Failed to move the generated file to the specified folder. The file will remain in the main Drive folder. Error: ' + e.toString());
            }
        }

        // 2. Find the template sheet in the current spreadsheet
        var templateSheet = spreadsheet.getSheetByName(templateName);
        if (!templateSheet) {
            throw new Error('Could not find the template sheet named "' + templateName + '" in the current spreadsheet.');
        }

        // 3. Copy the template sheet to the new spreadsheet
        var newSheet = templateSheet.copyTo(newSpreadsheet);
        newSheet.setName(CONFIG.SHEET_NAMES.NEW_PROPOSAL_SHEET); // Rename the copied sheet

        // 4. Delete the default "Sheet1" that is created automatically
        var defaultSheet = newSpreadsheet.getSheetByName('Sheet1');
        if (defaultSheet) {
            newSpreadsheet.deleteSheet(defaultSheet);
        }

        // 5. Replace placeholders in the new sheet
        for (var key in rowData) {
            newSheet.createTextFinder('{{' + key + '}}').replaceAllWith(rowData[key] || '');
        }

        // 6. Call the function to fill in the proposal data
        sudelioti_pasiulyma(spreadsheet, newSpreadsheet, rowData);
        
        // 7. Convert the filled spreadsheet to PDF without gridlines
        SpreadsheetApp.flush(); // Ensure all changes are saved

        var spreadsheetId = newSpreadsheet.getId();
        var sheetId = newSheet.getSheetId();
        var exportUrl = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?' +
                        'exportFormat=pdf&' +
                        'format=pdf&' +
                        'gid=' + sheetId + '&' +
                        'size=a4&' +           // A4 format
                        'portrait=true&' +     // Portrait orientation
                        'fitw=true&' +         // Fit to width
                        'sheetnames=false&' +
                        'printtitle=false&' +
                        'gridlines=false';     // Disable gridlines

        var response = UrlFetchApp.fetch(exportUrl, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } });
        var pdfBlob = response.getBlob().setName(newSpreadsheet.getName() + '.pdf');
        
        // Save the PDF file in the same specified folder
        var pdfTargetFolder;
        if (CONFIG.SHEET_NAMES.GENERATED_FILES_FOLDER_ID && CONFIG.SHEET_NAMES.GENERATED_FILES_FOLDER_ID !== 'JUSU_ARCHYVO_APLANKO_ID') {
            try {
                pdfTargetFolder = DriveApp.getFolderById(CONFIG.SHEET_NAMES.GENERATED_FILES_FOLDER_ID);
            } catch (e) {
                Logger.log('WARNING: Could not find the specified PDF folder. The PDF will be saved in the main Drive folder. Error: ' + e.toString());
                pdfTargetFolder = DriveApp.getRootFolder();
            }
        } else {
            pdfTargetFolder = DriveApp.getRootFolder(); // If no ID is specified, save in the root folder
        }
        var pdfFile = pdfTargetFolder.createFile(pdfBlob);
        
        // 8. Save the PDF link in the 'pasiulymu_nuorodos' column
        var pdfUrl = pdfFile.getUrl();
        var linksColumnName = 'pasiulymu_nuorodos';
        var linksColIndex = headers.map(h => h.toLowerCase()).indexOf(linksColumnName.toLowerCase());

        if (linksColIndex !== -1) {
            var creationDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
            var linkCell = sheet.getRange(rowIndex + 1, linksColIndex + 1);
            var existingLinks = linkCell.getValue().toString();
            var newLinksValue;
            if (existingLinks) {
                newLinksValue = existingLinks + '\n' + pdfUrl + '|' + creationDate; // Add the new link and date
            } else {
                newLinksValue = pdfUrl + '|' + creationDate;
            }
            linkCell.setValue(newLinksValue);
            Logger.log('Proposal link updated in column "' + linksColumnName + '" in row ' + (rowIndex + 1));
        } else {
            Logger.log('WARNING: Column "' + linksColumnName + '" not found. The proposal link was not saved.');
        }

        // 9. Prepare data for the email
        var emailConfigSheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.EMAIL_CONFIG);
        if (!emailConfigSheet) {
          throw new Error('Laiško konfigūracijos lapas "' + CONFIG.SHEET_NAMES.EMAIL_CONFIG + '" nerastas.');
        }
        // Nuskaitome temą ir turinį iš konfigūracijos lapo (antra eilutė)
        var subjectTemplate = emailConfigSheet.getRange('A2').getValue();
        var bodyTemplate = emailConfigSheet.getRange('B2').getValue();

        var recipientEmail = rowData['email'] || ''; // Make sure there is an 'email' column in the 'leads' sheet
        var subject = subjectTemplate;
        // Pakeičiame {{full_name}} į kliento vardą
        var body = bodyTemplate.replace('{{full_name}}', (rowData['full_name'] || ''));

        // 10. Create a Gmail draft
        var fromAlias = 'info@adenergy.lt';
        var aliases = GmailApp.getAliases();
        var options = {
            attachments: [pdfBlob],
            htmlBody: body.replace(/\n/g, '<br>') // Convert newlines to <br> for HTML format
        };

        if (aliases.includes(fromAlias)) {
            options.from = fromAlias;
        } else {
            Logger.log('WARNING: Alias "' + fromAlias + '" not found. The draft will be created using the default sender.');
        }

        var draft = GmailApp.createDraft(recipientEmail, subject, body, options);

        // 11. Return information to the client
        return {
            success: true,
            draftId: draft.getId(),
            message: 'Email draft successfully created.'
        };

    } catch (e) {
        Logger.log('Error generating document: ' + e.toString());
        throw new Error('Failed to generate document: ' + e.message);
    }
}
 
function sudelioti_pasiulyma(sourceSpreadsheet, targetSpreadsheet, rowData){
  try {
    Logger.log("sudelioti_pasiulyma: " + JSON.stringify(rowData));
    var targetSheet = targetSpreadsheet.getSheetByName(CONFIG.SHEET_NAMES.NEW_PROPOSAL_SHEET);
    if (!targetSheet) {
      throw new Error("Proposal sheet '" + CONFIG.SHEET_NAMES.NEW_PROPOSAL_SHEET + "' not found in the new spreadsheet.");
    }

    var calculatorSheetName = CONFIG.SHEET_NAMES.PROPOSAL_CALCULATION_PREFIX + 1;
    var calculatorSheet = sourceSpreadsheet.getSheetByName(calculatorSheetName);

    var today = new Date();
    var todayplus2weeks =new Date();
    todayplus2weeks.setDate(today.getDate() + 14);      
    var formattedDate_pasiulymodata = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var formattedDate_galiojaiki = Utilities.formatDate(todayplus2weeks, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var formattedDate_pasiulymonr = Utilities.formatDate(today, Session.getScriptTimeZone(), "MMdd"); 

    targetSheet.getRange('C2').setValue("Pasiūlymo data: "+formattedDate_pasiulymodata);
    targetSheet.getRange('C5').setValue("Galioja iki: "+formattedDate_galiojaiki);

    var pasiulymonr=calculatorSheet.getRange('C20').getValue();
    pasiulymonr = pasiulymonr + 1;
    calculatorSheet.getRange('C20').setValue(pasiulymonr);
    var pasiulymo_numeris='Nr.'+formattedDate_pasiulymonr+'-'+pasiulymonr
    
    targetSheet.getRange('B9').setValue(pasiulymo_numeris);
    targetSheet.getRange('A11').setValue("Pasiūlymo gavėjas:\n"+rowData["full_name"]);
    targetSpreadsheet.setName('Ad Energy pasiūlymas ' + pasiulymo_numeris);

    for (var i = 1; i <= rowData['pasiulymu_kiekis']; i++) {
      var calculatorSheetName = CONFIG.SHEET_NAMES.PROPOSAL_CALCULATION_PREFIX + i;
      var calculatorSheet = sourceSpreadsheet.getSheetByName(calculatorSheetName);
      if (!calculatorSheet) {
        Logger.log('WARNING: Calculation sheet "' + calculatorSheetName + '" not found. This proposal will be skipped.');
        continue;
      }
      
      calculatorSheet.getRange('C7').setValue(rowData["pasirinkite Kw"+i]);      
      calculatorSheet.getRange('C8').setValue(rowData["saules moduliai"+i]);      
      calculatorSheet.getRange('C9').setValue(rowData["konstrukcija"+i]);
       calculatorSheet.getRange('J29').setValue(rowData["nuolaida"+i]);

      Utilities.sleep(1000);
      SpreadsheetApp.flush(); // Svarbu: laukiame, kol formulės persiskaičiuos
        
      
      var sourceRange = calculatorSheet.getRange('G7:G21');
      var sourceRangeG6= calculatorSheet.getRange('G6');
      
      var targetColumn = String.fromCharCode('B'.charCodeAt(0) + i - 1); // B, C, D
      
      targetSheet.getRange(targetColumn + '13').setValue("Nr." + i + ". " + sourceRangeG6.getValue());
      targetSheet.getRange(targetColumn + '14:' + targetColumn + '28').setValues(sourceRange.getValues());
           
      
    }
    // tikrinam del nuolaidu
    var anyDiscountExists = false;
    for (var i = 1; i <= rowData['pasiulymu_kiekis']; i++) {
      if (rowData["nuolaida"+i] && Number(rowData["nuolaida"+i]) > 0) {
        anyDiscountExists = true;
        Logger.log("Rasta bent viena nuolaida, įterpiamos eilutės.");
        calculatorSheetName = CONFIG.SHEET_NAMES.PROPOSAL_CALCULATION_PREFIX + 1;
        calculatorSheet = sourceSpreadsheet.getSheetByName(calculatorSheetName);
        
        targetSheet.insertRowAfter(26); // Įterpiame eilutę nuolaidos reikšmei
        targetSheet.insertRowAfter(27); // Įterpiame eilutę nuolaidos pavadinimui
        targetSheet.getRange('A27').setValue(calculatorSheet.getRange('F29').getValue()).setFontColor("red");
        targetSheet.getRange('A28').setValue(calculatorSheet.getRange('F30').getValue());
        Logger.log("iterpti stulpeliu pavadinimai:"+calculatorSheet.getRange('F29').getValue()+" "+calculatorSheet.getRange('F30').getValue());
        break;
      }
    }

    // Jei bent viena nuolaida egzistuoja, įterpiame eilutes ir užpildome duomenis
    if (anyDiscountExists) {
      // Ciklas per pasiūlymus, kad užpildytume nuolaidų duomenis
      for (var i = 1; i <= rowData['pasiulymu_kiekis']; i++) {
        if (rowData["nuolaida"+i] && Number(rowData["nuolaida"+i]) > 0) {
          var calculatorSheetName = CONFIG.SHEET_NAMES.PROPOSAL_CALCULATION_PREFIX + i;
          var calculatorSheet = sourceSpreadsheet.getSheetByName(calculatorSheetName);
          var targetColumn = String.fromCharCode('B'.charCodeAt(0) + i - 1); // B, C, D

          if (calculatorSheet) {           
            targetSheet.getRange(targetColumn + '27').setValue(calculatorSheet.getRange('J29').getValue()).setFontColor("red");; // Nustatome vertę atitinkamame stulpelyje
            targetSheet.getRange(targetColumn + '28').setValue(calculatorSheet.getRange('J30').getValue());
            Logger.log("Pasiūlymui " + i + " pritaikyta nuolaida " + calculatorSheet.getRange('J29').getValue() + " ir " + calculatorSheet.getRange('J30').getValue() + " stulpelyje " + targetColumn);
          }
        }
      }
    }      
    
  }catch (e) {
        Logger.log('Error copying data: ' + e.toString());
        throw new Error('Error copying data: ' + e.message);
    }
}