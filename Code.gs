function doGet() {
  try {
    var spreadsheet = SpreadsheetApp.openById('1KFnjVbU4P0_YlUsKOIOhtIozduQwCSfr4v78J2o3GaM');
    if (!spreadsheet) {
      return HtmlService.createHtmlOutput('<p>Klaida: Nepavyko rasti lentelės su ID "1KFnjVbU4P0_YlUsKOIOhtIozduQwCSfr4v78J2o3GaM".</p>');
    }
    var sheet1 = spreadsheet.getSheetByName('leads');
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
                                        '. Patikrinkite "stulpeliai_rodyti" stulpelį config_pasiulymas lape ir įsitikinkite, kad atitinkami sunumeruoti stulpeliai (pvz., "' + 
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
    'A21': spreadsheet.getSheetByName('Kainos').getRange('A21').getValue().toString().trim(),
    'A23': spreadsheet.getSheetByName('Kainos').getRange('A23').getValue().toString().trim(),
    'A24': spreadsheet.getSheetByName('Kainos').getRange('A24').getValue().toString().trim(),
    'A38': spreadsheet.getSheetByName('Kainos').getRange('A38').getValue().toString().trim(),
    'A42': spreadsheet.getSheetByName('Kainos').getRange('A42').getValue().toString().trim()
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
      pasiulymoOptions.push(pasiulymoReiksmes[index] ? getCellValueOrFormula(pasiulymoReiksmes[index], 0) : []);
      formulaOptions.push(formulas[index] ? getCellValueOrFormula(formulas[index], 0) : []);
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
      group1Columns.push({ index: idx, displayName: displayNames[i], editable: isEditable[i], options: columnOptions[i], dateButton: hasDateButton[i], rowCount: rowCountsConfig[i], isDateTime: isDateTimeColumn[i], isDate: isDateColumn[i], hasDatePicker: hasDatePicker[i], pasiulymoOptions: pasiulymoOptions[i], formulaOptions: formulaOptions[i], columnName: columnNames[i] });
    } else if (positions[i] === 2) {
      group2Columns.push({ index: idx, displayName: displayNames[i], editable: isEditable[i], options: columnOptions[i], dateButton: hasDateButton[i], rowCount: rowCountsConfig[i], isDateTime: isDateTimeColumn[i], isDate: isDateColumn[i], hasDatePicker: hasDatePicker[i], pasiulymoOptions: pasiulymoOptions[i], formulaOptions: formulaOptions[i], columnName: columnNames[i] });
    }
  });
  
  // Process columns for proposal modal
  var proposalColumnIndices = [];
  var proposalDisplayNames = [];
  var proposalIsEditable = [];
  var proposalColumnOptions = [];
  var proposalColumnNames = [];
  var allProposalColumns = [];

  proposal.selectedColumns.forEach(function(colName, index) {
    var baseColName = colName.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase();
    var popupOptions = [];
    if (proposal.popupColumns[index]) {
      popupOptions = getCellValueOrFormula(proposal.popupColumns[index], 0);
    }

    // Find all numbered versions of the column (e.g., "pasirinkite kw1", "pasirinkite kw2", etc.)
    for (var p = 1; p <= 3; p++) {
      var numberedColName = colName + p;
      var colIndex = headers.indexOf(numberedColName.toLowerCase());
      if (colIndex !== -1) {
        allProposalColumns.push({
          index: colIndex,
          displayName: colName, // Base name for display
          editable: proposal.editableColumns[index] === 'x',
          options: popupOptions,
          columnName: baseColName + p // Unique name, e.g., pasirinkite_kw1
        });
      }
    }
  });
  
  if (allProposalColumns.length === 0) {
    return HtmlService.createHtmlOutput('<p>Klaida: Nerasta tinkamų stulpelių leads antraštėse pagal config_pasiulymas. Patikrinkite "stulpeliai_rodyti" stulpelį config_pasiulymas lape.</p>');
  }

  // Base structure for the modal (uses base names)
  var proposalColumns = proposal.selectedColumns.map((colName, i) => ({
    displayName: colName,
    editable: proposal.editableColumns[i] === 'x',
    options: getCellValueOrFormula(proposal.popupColumns[i], 0),
    dateButton: false,
    rowCount: 1,
    isDateTime: false,
    isDate: false,
    hasDatePicker: false,
    pasiulymoOptions: [],
    formulaOptions: [],
    columnName: colName.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase()
  }));
  
  // Get pasiulymu_kiekis options and index
  var pasiulymuKiekisIndex = headers.indexOf('pasiulymu_kiekis');
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
  html.pasiulymuKiekisOptions = pasiulymuKiekisOptions;
  html.pasiulymuKiekisIndex = pasiulymuKiekisIndex;
  html.kainosValues = kainosValues;
  // Surandame ir perduodame 'id' stulpelio indeksą
  var idIndex = headers.indexOf('id');
  html.idIndex = idIndex;

  return html.evaluate().setTitle('Facebook Leads Duomenys').setWidth(1000).setHeight(600);
  } catch (e) {
    Logger.log('Klaida vykdant doGet: ' + e.toString());
    return HtmlService.createHtmlOutput('<p>Įvyko kritinė klaida: ' + e.toString() + '</p>');
  }
}

function updateSheet1(rowIndex, updates) {
  var spreadsheet = SpreadsheetApp.openById('1KFnjVbU4P0_YlUsKOIOhtIozduQwCSfr4v78J2o3GaM');
  if (!spreadsheet) {
    Logger.log('Error: Nepavyko rasti lentelės su ID "1KFnjVbU4P0_YlUsKOIOhtIozduQwCSfr4v78J2o3GaM"');
    throw new Error('Nepavyko rasti lentelės su ID "1KFnjVbU4P0_YlUsKOIOhtIozduQwCSfr4v78J2o3GaM"');
  }
  var sheet = spreadsheet.getSheetByName('leads');
  if (!sheet) {
    Logger.log('Error: Lenta "leads" nerasta');
    throw new Error('Lenta "leads" nerasta');
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

function generateProposalDocument(uniqueId) {
    try {
        var spreadsheet = SpreadsheetApp.openById('1KFnjVbU4P0_YlUsKOIOhtIozduQwCSfr4v78J2o3GaM');
        var sheet = spreadsheet.getSheetByName('leads');
        var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var allData = sheet.getDataRange().getValues();

        // Surandame stulpelio, kurį naudosime kaip unikalų ID, indeksą (pvz., 'id')
        var uniqueIdColIndex = headers.map(h => h.toLowerCase()).indexOf('id');
        if (uniqueIdColIndex === -1) {
            throw new Error("Nepavyko rasti 'id' stulpelio, kuris naudojamas kaip unikalus identifikatorius.");
        }

        // Surandame eilutę, atitinkančią unikalų ID
        var rowIndex = -1;
        var data;
        allData.forEach((row, i) => {
            if (i > 0 && row[uniqueIdColIndex].toString() === uniqueId) {
                rowIndex = i; // i yra indeksas allData masyve (prasideda nuo 0)
                data = row;
            }
        });

        if (!data) {
            throw new Error('Nepavyko rasti eilutės su ID: ' + uniqueId);
        }

        var rowData = {};
        headers.forEach((header, i) => {
            rowData[header] = data[i];
        });

        // Patikimesnis būdas rasti 'pasiulymu_kiekis' reikšmę, nepriklausomai nuo raidžių dydžio
        var proposalCountKey = Object.keys(rowData).find(key => key.toLowerCase() === 'pasiulymu_kiekis');
        var proposalCount = 1; // Numatytasis, jei nerandama arba reikšmė netinkama
        if (proposalCountKey && rowData[proposalCountKey]) {
            var countValue = parseInt(rowData[proposalCountKey], 10);
            if (!isNaN(countValue) && countValue > 0) {
                proposalCount = countValue;
            }
        }
        var templateName = 'template_pasiulymas' + proposalCount;

        Logger.log('Generuojamas pasiūlymas eilutei: ' + (rowIndex + 1) + ' naudojant šabloną: ' + templateName);
        Logger.log('data: '+data)
        Logger.log('rowData: ' + JSON.stringify(rowData, null, 2));
        
        // 1. Sukuriame naują Google Sheets failą
        var newSpreadsheet = SpreadsheetApp.create('Pasiūlymas ' + (rowData['Vardas'] || uniqueId));
        
        // 2. Surandame šablono lapą (Sheet) dabartinėje lentelėje
        var templateSheet = spreadsheet.getSheetByName(templateName);
        if (!templateSheet) {
            throw new Error('Nepavyko rasti šablono lapo (Sheet) pavadinimu "' + templateName + '" dabartinėje lentelėje.');
        }

        // 3. Nukopijuojame šablono lapą į naują lentelę
        var newSheet = templateSheet.copyTo(newSpreadsheet);
        newSheet.setName('Pasiūlymas'); // Pervadiname nukopijuotą lapą

        // 4. Ištriname numatytąjį "Sheet1" lapą, kuris sukuriamas automatiškai
        var defaultSheet = newSpreadsheet.getSheetByName('Sheet1');
        if (defaultSheet) {
            newSpreadsheet.deleteSheet(defaultSheet);
        }

        // 5. Pakeičiame žymeklius naujame lape
        for (var key in rowData) {
            newSheet.createTextFinder('{{' + key + '}}').replaceAllWith(rowData[key] || '');
        }
    
        return { url: newSpreadsheet.getUrl() };
    } catch (e) {
        Logger.log('Klaida generuojant dokumentą: ' + e.toString());
        throw new Error('Nepavyko sugeneruoti dokumento: ' + e.message);
    }

    selectFromDropdown('1KFnjVbU4P0_YlUsKOIOhtIozduQwCSfr4v78J2o3GaM','pasiulymas',"C7","8")
}
 
//function sudelioti_pasiulyma(variantas){


//}

function selectFromDropdown(spreadsheetId, sheetName, range, desiredValue) {
  try {
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('Lapas "' + sheetName + '" nerastas lentelėje su ID ' + spreadsheetId);
    }

    var cell = sheet.getRange(range);
    cell.setValue(desiredValue);

    Logger.log('Sėkmingai nustatyta reikšmė "' + desiredValue + '" langelyje ' + range + ' lape "' + sheetName + '".');
    return { success: true, message: 'Reikšmė nustatyta.' };
  } catch (e) {
    Logger.log('Klaida vykdant selectFromDropdown: ' + e.toString());
    throw new Error('Nepavyko nustatyti reikšmės: ' + e.message);
  }
}