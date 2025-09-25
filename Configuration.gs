/**
 * Įkelia visą konfigūraciją iš 'configuration' ir 'config_pasiulymas' lapų.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet Aktyvi skaičiuoklė.
 * @returns {Object} Konfigūracijos objektas.
 */
function loadConfiguration(spreadsheet) {
  const sheet2 = spreadsheet.getSheetByName('configuration');
  const sheetProposal = spreadsheet.getSheetByName('config_pasiulymas');

  if (!sheet2 || !sheetProposal) {
    throw new Error('Trūksta vieno ar daugiau būtinų lapų ("configuration" arba "config_pasiulymas").');
  }

  const sheet2Headers = sheet2.getRange(1, 1, 1, sheet2.getLastColumn()).getValues()[0];

  const requiredColumns = [
    'stulpeliai_rodyti', 'stulpeliai_rodyti_map', 'issokusi_redaguoti',
    'issokusi_data_mygtukas', 'issokusi_laukelio_eilutes', 'stulpeliai_YYYY-MM-DD HH:MM',
    'Stulpeliai', 'pasiulymo_reiksmes', 'formule'
  ];

  for (const colName of requiredColumns) {
    if (sheet2Headers.indexOf(colName) === -1) {
      throw new Error(`Klaida: Stulpelis "${colName}" nerastas configuration lape.`);
    }
  }

  const columnToShowIndex = sheet2Headers.indexOf('stulpeliai_rodyti');
  const columnMapIndex = sheet2Headers.indexOf('stulpeliai_rodyti_map');
  const columnEditableIndex = sheet2Headers.indexOf('issokusi_redaguoti');
  const columnOptionsIndex = sheet2Headers.indexOf('issokusi_reiksmes');
  const columnDateButtonIndex = sheet2Headers.indexOf('issokusi_data_mygtukas');
  const columnRowsIndex = sheet2Headers.indexOf('issokusi_laukelio_eilutes');
  const columnDateTimeColumnIndex = sheet2Headers.indexOf('stulpeliai_YYYY-MM-DD HH:MM');
  const columnDateColumnIndex = sheet2Headers.indexOf('stulpeliai_YYYY-MM-DD');
  const columnDatePickerIndex = sheet2Headers.indexOf('issokusi_data_pasirnkti_data');
  const columnPositionIndex = sheet2Headers.indexOf('Stulpeliai');
  const columnPasiulymoReiksmesIndex = sheet2Headers.indexOf('pasiulymo_reiksmes');
  const columnFormulaIndex = sheet2Headers.indexOf('formule');

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
  const columnPositions = sheet2.getRange(2, columnPositionIndex + 1, maxRows, 1).getValues().flat().map(val => val ? parseInt(val) : 0);
  const pasiulymoReiksmes = sheet2.getRange(2, columnPasiulymoReiksmesIndex + 1, maxRows, 1).getValues().flat();
  const formulas = sheet2.getRange(2, columnFormulaIndex + 1, maxRows, 1).getValues().flat();

  // Proposal config
  const sheetProposalHeaders = sheetProposal.getRange(1, 1, 1, sheetProposal.getLastColumn()).getValues()[0];
  const requiredProposalColumns = ['stulpeliai_rodyti', 'issokusi_redaguoti', 'issokusi_pupop'];

  for (const colName of requiredProposalColumns) {
    if (sheetProposalHeaders.indexOf(colName) === -1) {
      throw new Error(`Klaida: Stulpelis "${colName}" nerastas config_pasiulymas lape.`);
    }
  }

  const proposalColumnToShowIndex = sheetProposalHeaders.indexOf('stulpeliai_rodyti');
  const proposalColumnEditableIndex = sheetProposalHeaders.indexOf('issokusi_redaguoti');
  const proposalColumnPopupIndex = sheetProposalHeaders.indexOf('issokusi_pupop');

  const maxProposalRows = sheetProposal.getLastRow() - 1;
  const proposalSelectedColumns = sheetProposal.getRange(2, proposalColumnToShowIndex + 1, maxProposalRows, 1).getValues().flat().filter(String).map(col => col.trim());
  const proposalEditableColumns = sheetProposal.getRange(2, proposalColumnEditableIndex + 1, maxProposalRows, 1).getValues().flat();
  const proposalPopupColumns = sheetProposal.getRange(2, proposalColumnPopupIndex + 1, maxProposalRows, 1).getValues().flat();

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
    formulas,
    proposal: {
      selectedColumns: proposalSelectedColumns,
      editableColumns: proposalEditableColumns,
      popupColumns: proposalPopupColumns
    }
  };
}