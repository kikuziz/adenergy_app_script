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

  // Read all data from the configuration sheet at once
  const configData = sheet2.getRange(2, 1, sheet2.getLastRow() - 1, sheet2.getLastColumn()).getValues();

  // Helper function to get a column's data from the pre-fetched array
  const getColumnData = (headerName) => {
    const index = sheet2Headers.indexOf(headerName);
    if (index === -1) return new Array(configData.length).fill('');
    return configData.map(row => row[index]);
  };

  const selectedColumns = getColumnData('stulpeliai_rodyti').filter(String).map(col => col.toString().trim());
  const mappedNames = getColumnData('stulpeliai_rodyti_map');
  const editableColumns = getColumnData('redaguoti');
  const optionsColumns = getColumnData('issokusi_reiksmes');
  const dateButtonColumns = getColumnData('issokusi_data_mygtukas');
  const rowCounts = getColumnData('issokusi_laukelio_eilutes');
  const dateTimeColumns = getColumnData('stulpeliai_YYYY-MM-DD HH:MM');
  const dateColumns = getColumnData('stulpeliai_YYYY-MM-DD');
  const datePickerColumns = getColumnData('issokusi_data_pasirnkti_data');
  const columnPositions = getColumnData('Stulpeliai').map(val => (val !== "" && !isNaN(val)) ? parseInt(val) : 1);
  const pasiulymoReiksmes = getColumnData('pasiulymo_reiksmes');

  // Proposal config
  const sheetProposalHeaders = sheetProposal.getRange(1, 1, 1, sheetProposal.getLastColumn()).getValues()[0];
  const requiredProposalColumns = ['pasiulymas_stulpeliai_rodyti', 'redaguoti', 'dropdown_link', 'autoupdate_to_pasiulymas'];

  for (const colName of requiredProposalColumns) {
    if (sheetProposalHeaders.indexOf(colName) === -1) {
      // This can be a simple warning, no need to throw an error if some are optional
      Logger.log(`Warning: Column "${colName}" not found in sheet "${CONFIG.SHEET_NAMES.PROPOSAL_CONFIG}".`);
    }
  }

  // Read all data from the proposal config sheet at once
  const proposalConfigData = sheetProposal.getLastRow() > 1 ? sheetProposal.getRange(2, 1, sheetProposal.getLastRow() - 1, sheetProposal.getLastColumn()).getValues() : [];

  const getProposalColumnData = (headerName) => {
    const index = sheetProposalHeaders.indexOf(headerName);
    if (index === -1) return new Array(proposalConfigData.length).fill('');
    return proposalConfigData.map(row => row[index]);
  };

  const proposalSelectedColumns = getProposalColumnData('pasiulymas_stulpeliai_rodyti').filter(String).map(col => col.toString().trim());
  const proposalEditableColumns = getProposalColumnData('redaguoti');
  const proposalDropdownLinkColumns = getProposalColumnData('dropdown_link');
  const proposalAutoUpdateColumns = getProposalColumnData('autoupdate_to_pasiulymas');

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