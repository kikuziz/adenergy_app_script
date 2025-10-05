/**
 * Įkelia visą konfigūraciją iš 'configuration' ir 'config_pasiulymas' lapų.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Aktyvi skaičiuoklė.
 * @returns {Object} Konfigūracijos objektas.
 */
function loadConfiguration(spreadsheet) {
  const sheet2 = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.CONFIGURATION);
  const sheetProposal = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.PROPOSAL_CONFIG);
  const sheetProposalAutoDisplay = spreadsheet.getSheetByName(CONFIG.SHEET_NAMES.PROPOSAL_AUTO_DISPLAY_CONFIG);

  if (!sheet2 || !sheetProposal || !sheetProposalAutoDisplay) {
    throw new Error('Trūksta vieno ar daugiau būtinų konfigūracijos lapų ("' + CONFIG.SHEET_NAMES.CONFIGURATION + '", "' + CONFIG.SHEET_NAMES.PROPOSAL_CONFIG + '", arba "' + CONFIG.SHEET_NAMES.PROPOSAL_AUTO_DISPLAY_CONFIG + '").');
  }

  const sheet2Headers = sheet2.getRange(1, 1, 1, sheet2.getLastColumn()).getValues()[0];

  const requiredColumns = [
    'stulpeliai_rodyti', 'stulpeliai_rodyti_map', 'redaguoti',
    'issokusi_data_mygtukas', 'issokusi_laukelio_eilutes', 'stulpeliai_YYYY-MM-DD HH:MM',
    'Stulpeliai', 'pasiulymo_reiksmes'
  ];

  for (const colName of requiredColumns) {
    if (sheet2Headers.indexOf(colName) === -1) {
      throw new Error(`Klaida: Stulpelis "${colName}" nerastas ${CONFIG.SHEET_NAMES.CONFIGURATION} lape.`);
    }
  }

  const columnToShowIndex = sheet2Headers.indexOf('stulpeliai_rodyti');
  const columnMapIndex = sheet2Headers.indexOf('stulpeliai_rodyti_map');
  const columnEditableIndex = sheet2Headers.indexOf('redaguoti');
  const columnOptionsIndex = sheet2Headers.indexOf('issokusi_reiksmes');
  const columnDateButtonIndex = sheet2Headers.indexOf('issokusi_data_mygtukas');
  const columnRowsIndex = sheet2Headers.indexOf('issokusi_laukelio_eilutes');
  const columnDateTimeColumnIndex = sheet2Headers.indexOf('stulpeliai_YYYY-MM-DD HH:MM');
  const columnDateColumnIndex = sheet2Headers.indexOf('stulpeliai_YYYY-MM-DD');
  const columnDatePickerIndex = sheet2Headers.indexOf('issokusi_data_pasirnkti_data');
  const columnPositionIndex = sheet2Headers.indexOf('Stulpeliai');
  const columnPasiulymoReiksmesIndex = sheet2Headers.indexOf('pasiulymo_reiksmes');

  const maxRows = sheet2.getLastRow() - 1;
  const selectedColumns = sheet2.getRange(2, columnToShowIndex + 1, maxRows, 1).getValues().flat().filter(String).map(col => col.trim());
  const mappedNames = sheet2.getRange(2, columnMapIndex + 1, maxRows, 1).getValues().flat();
  const editableColumns = sheet2.getRange(2, columnEditableIndex + 1, maxRows, 1).getValues().flat();
  const optionsColumns = columnOptionsIndex !== -1 ? sheet2.getRange(2, columnOptionsIndex + 1, maxRows, 1).getValues().flat() : new Array(maxRows).fill('');
  const dateButtonColumns = sheet2.getRange(2, columnDateButtonIndex + 1, maxRows, 1).getValues().flat();
  const rowCounts = columnRowsIndex !== -1 ? sheet2.getRange(2, columnRowsIndex + 1, maxRows, 1).getValues().flat() : new Array(maxRows).fill('');
  const dateTimeColumns = sheet2.getRange(2, columnDateTimeColumnIndex + 1, maxRows, 1).getValues().flat();
  const dateColumns = columnDateColumnIndex !== -1 ? sheet2.getRange(2, columnDateColumnIndex + 1, maxRows, 1).getValues().flat() : new Array(maxRows).fill('');
  const datePickerColumns = columnDatePickerIndex !== -1 ? sheet2.getRange(2, columnDatePickerIndex + 1, maxRows, 1).getValues().flat() : new Array(maxRows).fill('');
  const columnPositions = sheet2.getRange(2, columnPositionIndex + 1, maxRows, 1).getValues().flat().map(val => (val !== "" && !isNaN(val)) ? parseInt(val) : 1);
  const pasiulymoReiksmes = sheet2.getRange(2, columnPasiulymoReiksmesIndex + 1, maxRows, 1).getValues().flat();

  // Proposal config
  const sheetProposalHeaders = sheetProposal.getRange(1, 1, 1, sheetProposal.getLastColumn()).getValues()[0];
  const requiredProposalColumns = ['pasiulymas_stulpeliai_rodyti', 'redaguoti', 'dropdown_link', 'autoupdate_to_pasiulymas'];

  for (const colName of requiredProposalColumns) {
    if (sheetProposalHeaders.indexOf(colName) === -1) {
      Logger.log('DĖMESIO: Stulpelis "' + colName + '" nerastas ' + CONFIG.SHEET_NAMES.PROPOSAL_CONFIG + ' lape, bet tęsiama toliau.');
    }
  }

  const proposalColumnToShowIndex = sheetProposalHeaders.indexOf('pasiulymas_stulpeliai_rodyti');
  const proposalColumnEditableIndex = sheetProposalHeaders.indexOf('redaguoti');
  const proposalColumnDropdownLinkIndex = sheetProposalHeaders.indexOf('dropdown_link');
  const proposalColumnAutoUpdateIndex = sheetProposalHeaders.indexOf('autoupdate_to_pasiulymas');

  const maxProposalRows = sheetProposal.getLastRow() - 1;
  const proposalSelectedColumns = sheetProposal.getRange(2, proposalColumnToShowIndex + 1, maxProposalRows, 1).getValues().flat().filter(String).map(col => col.trim());
  const proposalEditableColumns = sheetProposal.getRange(2, proposalColumnEditableIndex + 1, maxProposalRows, 1).getValues().flat();
  const proposalDropdownLinkColumns = sheetProposal.getRange(2, proposalColumnDropdownLinkIndex + 1, maxProposalRows, 1).getValues().flat();
  const proposalAutoUpdateColumns = proposalColumnAutoUpdateIndex !== -1 ? sheetProposal.getRange(2, proposalColumnAutoUpdateIndex + 1, maxProposalRows, 1).getValues().flat() : new Array(maxProposalRows).fill('');

  // Load auto-display config
  const autoDisplayHeaders = sheetProposalAutoDisplay.getRange(1, 1, 1, sheetProposalAutoDisplay.getLastColumn()).getValues()[0];
  const pavadinimaiIndex = autoDisplayHeaders.indexOf('pasiulymas_pavadinimas');
  const reiksmeIndex = autoDisplayHeaders.indexOf('pasiulymas_reiksme');

  if (pavadinimaiIndex === -1 || reiksmeIndex === -1) {
    throw new Error('Trūksta būtinų stulpelių "pasiulymas_pavadinimas" arba "pasiulymas_reiksme" lape ' + CONFIG.SHEET_NAMES.PROPOSAL_AUTO_DISPLAY_CONFIG);
  }

  const autoDisplayData = sheetProposalAutoDisplay.getLastRow() > 1 ? sheetProposalAutoDisplay.getRange(2, 1, sheetProposalAutoDisplay.getLastRow() - 1, sheetProposalAutoDisplay.getLastColumn()).getValues() : [];

  const autoDisplayConfig = autoDisplayData.map(row => ({
    pavadinimasRef: row[pavadinimaiIndex], // e.g., 'pasiulymas!B10'
    reiksmeRef: row[reiksmeIndex]      // e.g., 'pasiulymas!C10'
  })).filter(item => item.pavadinimasRef && item.reiksmeRef);

  return {
    selectedColumns,
    mappedNames,
    editableColumns,
    optionsColumns,
    dateButtonColumns,
    rowCounts,
    dateTimeColumns,
    dateColumns,
    datePickerColumns,
    columnPositions,
    pasiulymoReiksmes,
    proposal: {
      selectedColumns: proposalSelectedColumns,
      editableColumns: proposalEditableColumns,
      dropdownLinkColumns: proposalDropdownLinkColumns,
      autoUpdateColumns: proposalAutoUpdateColumns,
      autoDisplay: autoDisplayConfig
    },
  };
}