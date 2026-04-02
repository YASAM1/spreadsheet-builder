import express from 'express';
import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const DEFAULT_PORT = process.env.PORT || 3000;

export const TEMPLATE_PATH = path.join(
  __dirname,
  'template',
  '(API) 2025.07.30.Inventory-MediationSpreadsheetMaster.xlsx'
);

export const SHEET_NAME = 'Prop Distr - KB Preferred - Use';

export const SECTION_DEFINITIONS = [
  {
    key: 'Real Property',
    headerText: 'Real Property',
    aliases: ['Real Properties']
  },
  {
    key: 'Bank Accounts',
    headerText: 'Bank Accounts',
    aliases: ['Cash and Accounts with Financial Institutions']
  },
  {
    key: 'Business Interests',
    headerText: 'Business Interests',
    aliases: ['Business Interest', 'Closely Held Business Interests']
  },
  {
    key: 'Retirement Benefits',
    headerText: 'Retirement Benefits',
    aliases: ['Defined Contribution Plans']
  },
  {
    key: 'Stock Options',
    headerText: 'Stock Options',
    aliases: ['Stock Option']
  },
  {
    key: 'Personal Property',
    headerText: 'Personal Property (Furniture, Electronics, Firearms, Jewelry, etc)',
    aliases: [
      'Miscellaneous Sporting Goods And Firearms',
      'Money Owed To Me Or My Spouse',
      'Personal Property (Furniture, Electronics, Firearms, Jewelry, etc)'
    ]
  },
  {
    key: 'Life Insurance / Annuities',
    headerText: 'Life Insurance / Annuities',
    aliases: ['Insurance And Annuities', 'Life Insurance']
  },
  {
    key: 'Vehicles',
    headerText: 'Vehicles',
    aliases: ['Motor Vehicles, Boats, Airplanes, Cycles, etc']
  },
  {
    key: 'Credit Cards',
    headerText: 'Credit Cards',
    aliases: ['Credit Card']
  },
  {
    key: 'Bonuses',
    headerText: 'Bonuses',
    aliases: ['Bonus']
  },
  {
    key: 'Other Debts & Liabilities',
    headerText: 'Other Debts & Liabilities',
    aliases: ['Other Debt & Liabilities']
  },
  {
    key: 'Federal Income Taxes',
    headerText: 'Federal Income Taxes',
    aliases: ['Federal Income Tax']
  },
  {
    key: "Attorney's & Expert Fees",
    headerText: "Attorney's & Expert Fees",
    aliases: ['Attorney Fees', 'Expert Fees']
  },
  {
    key: "Spouse 1's Separate Property",
    headerText: "Spouse 1's Separate Property",
    aliases: ['Spouse 1’s Separate Property']
  },
  {
    key: "Spouse 2's Separate Property",
    headerText: "Spouse 2's Separate Property",
    aliases: ['Spouse 2’s Separate Property']
  },
  {
    key: "Children's Property",
    headerText: "Children's Property",
    aliases: ['Children’s Property']
  }
];

const SUMMARY_HEADERS = [
  'Sub Totals',
  '$ Difference between Spouse 1 and Spouse 2',
  '$ Difference between Spouse 1 and Spouse 2 ÷ 2',
  '% Distribution before Equalization',
  'Amt needed to Equalize (Use figure from Cell "$ Difference between Spouse 2 and Spouse 2 ÷ 2" and reverse +/-)',
  'Total'
];

const EXTRA_HEADING_TEXTS = [
  'Asset',
  'Assets',
  'Liability',
  'Community Estate',
  'Community Assets',
  'Amount To Petitioner',
  'Amount To Respondent',
  'In Possession Of Petitioner',
  'In Possession Of Respondent',
  'Health Savings Accounts',
  'Publicly Traded Stocks Bonds And Other Securities'
];

const DEBT_SECTIONS = new Set([
  'Credit Cards',
  'Other Debts & Liabilities',
  'Federal Income Taxes',
  "Attorney's & Expert Fees"
]);

const ROUTING_RULES = [
  { key: 'Credit Cards', match: isCreditCardLikeText },
  { key: 'Federal Income Taxes', match: isFederalIncomeTaxLikeText },
  { key: "Attorney's & Expert Fees", match: isAttorneyFeeLikeText },
  { key: 'Business Interests', match: isBusinessInterestLikeText },
  { key: 'Stock Options', match: isStockOptionLikeText },
  { key: 'Bonuses', match: isBonusLikeText }
];

const SECTION_FORMULA_PATTERNS = new Map(
  SECTION_DEFINITIONS.map(({ key }) => {
    const isSeparatePropertySection =
      key === "Spouse 1's Separate Property" ||
      key === "Spouse 2's Separate Property" ||
      key === "Children's Property";

    return [
      key,
      isSeparatePropertySection
        ? [
            { column: 3, formula: 'A{row}+B{row}' }
          ]
        : [
            { column: 3, formula: 'A{row}+B{row}' },
            { column: 5, formula: 'C{row}*0.5' },
            { column: 6, formula: 'C{row}-E{row}' }
          ]
    ];
  })
);

const SECTION_HEADER_BY_KEY = new Map();
const SECTION_KEY_BY_NORMALIZED_NAME = new Map();

for (const definition of SECTION_DEFINITIONS) {
  SECTION_HEADER_BY_KEY.set(definition.key, definition.headerText);

  const aliases = [definition.key, definition.headerText, ...(definition.aliases || [])];
  for (const alias of aliases) {
    SECTION_KEY_BY_NORMALIZED_NAME.set(normalizeHeadingText(alias), definition.key);
  }
}

const PROTECTED_HEADERS = [
  ...SECTION_DEFINITIONS.map(({ headerText }) => headerText),
  ...SUMMARY_HEADERS
];

const KNOWN_HEADING_TEXTS = new Set(
  [...PROTECTED_HEADERS, ...EXTRA_HEADING_TEXTS, ...SECTION_DEFINITIONS.flatMap(({ aliases = [] }) => aliases)].map(
    normalizeHeadingText
  )
);

const CELL_REFERENCE_REGEX = /(\$?)([A-Z]{1,3})(\$?)(\d+)/g;

function normalizeApostrophes(value) {
  if (!value || typeof value !== 'string') {
    return value || '';
  }

  return value.replace(/[\u2018\u2019\u201A\u201B]/g, "'");
}

function normalizeHeadingText(value) {
  return normalizeApostrophes(String(value || ''))
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function getSectionHeader(sectionKey) {
  const canonicalKey = resolveSectionKey(sectionKey);
  return canonicalKey ? SECTION_HEADER_BY_KEY.get(canonicalKey) : null;
}

function resolveSectionKey(sectionKey) {
  if (!sectionKey) {
    return null;
  }

  return SECTION_KEY_BY_NORMALIZED_NAME.get(normalizeHeadingText(sectionKey)) || null;
}

export function extractSections(body) {
  if (Array.isArray(body)) {
    if (body.length > 0 && body[0].sections) {
      return body[0].sections;
    }
    if (body.length > 0 && body[0].json && body[0].json.sections) {
      return body[0].json.sections;
    }
    return null;
  }

  if (body.sections) {
    return body.sections;
  }

  if (body.json && body.json.sections) {
    return body.json.sections;
  }

  return null;
}

function getCellStringValue(cell) {
  try {
    if (!cell || cell.value === null || cell.value === undefined) {
      return '';
    }

    if (typeof cell.value === 'string') {
      return cell.value;
    }

    if (typeof cell.value === 'number') {
      return String(cell.value);
    }

    if (typeof cell.value === 'object') {
      if (cell.value.error) {
        return '';
      }

      if (cell.value.richText) {
        return cell.value.richText.map((item) => item.text || '').join('');
      }

      if (cell.value.formula && cell.value.result !== undefined) {
        if (typeof cell.value.result === 'object' && cell.value.result?.error) {
          return '';
        }
        return String(cell.value.result || '');
      }

      if (cell.value.text) {
        return cell.value.text;
      }

      if (cell.value instanceof Date) {
        return cell.value.toISOString();
      }
    }

    return String(cell.value || '');
  } catch {
    return '';
  }
}

function getCellFormula(cell) {
  if (!cell) {
    return null;
  }

  if (cell.value && typeof cell.value === 'object') {
    if (typeof cell.value.formula === 'string' && cell.value.formula.trim()) {
      return cell.value.formula;
    }

    if (typeof cell.value.sharedFormula === 'string' && cell.worksheet) {
      const masterCell = cell.worksheet.getCell(cell.value.sharedFormula);
      if (masterCell && masterCell !== cell) {
        const masterFormula = getCellFormula(masterCell);
        if (masterFormula) {
          return shiftFormulaRows(masterFormula, masterCell.row, cell.row);
        }
      }
    }
  }

  if (cell.model) {
    if (typeof cell.model.formula === 'string' && cell.model.formula.trim()) {
      return cell.model.formula;
    }

    if (typeof cell.model.sharedFormula === 'string' && cell.worksheet) {
      const masterCell = cell.worksheet.getCell(cell.model.sharedFormula);
      if (masterCell && masterCell !== cell) {
        const masterFormula = getCellFormula(masterCell);
        if (masterFormula) {
          return shiftFormulaRows(masterFormula, masterCell.row, cell.row);
        }
      }
    }
  }

  try {
    if (typeof cell.formula === 'string' && cell.formula.trim()) {
      return cell.formula;
    }
  } catch {
    return null;
  }

  return null;
}

function isKnownHeader(value) {
  if (!value || typeof value !== 'string') {
    return false;
  }

  return KNOWN_HEADING_TEXTS.has(normalizeHeadingText(value));
}

function findHeaderRow(worksheet, headerText) {
  const targetText = normalizeHeadingText(headerText);
  let foundRow = null;

  worksheet.eachRow((row, rowNumber) => {
    const cellValue = getCellStringValue(row.getCell(4));
    if (normalizeHeadingText(cellValue) === targetText) {
      foundRow = rowNumber;
    }
  });

  return foundRow;
}

function findSectionEnd(worksheet, startRow) {
  let endRow = startRow;
  const lastRow = worksheet.lastRow ? worksheet.lastRow.number : startRow;

  for (let rowNumber = startRow + 1; rowNumber <= lastRow + 1; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const cellValue = getCellStringValue(row.getCell(4));

    if (isKnownHeader(cellValue)) {
      break;
    }

    endRow = rowNumber;
  }

  return endRow;
}

function copyCellStyle(sourceCell, targetCell) {
  if (sourceCell.font) {
    targetCell.font = { ...sourceCell.font };
  }
  if (sourceCell.fill) {
    targetCell.fill = JSON.parse(JSON.stringify(sourceCell.fill));
  }
  if (sourceCell.border) {
    targetCell.border = JSON.parse(JSON.stringify(sourceCell.border));
  }
  if (sourceCell.alignment) {
    targetCell.alignment = { ...sourceCell.alignment };
  }
  if (sourceCell.numFmt) {
    targetCell.numFmt = sourceCell.numFmt;
  }
}

function clearCellValue(cell) {
  const formula = getCellFormula(cell);
  if (formula || cell.sharedFormula) {
    return;
  }

  cell.value = null;
}

function isProtectedRow(worksheet, rowNumber) {
  const row = worksheet.getRow(rowNumber);
  return isKnownHeader(getCellStringValue(row.getCell(4)));
}

function getWritableRows(worksheet, sectionStart, sectionEnd) {
  const rows = [];

  for (let rowNumber = sectionStart; rowNumber <= sectionEnd; rowNumber += 1) {
    if (!isProtectedRow(worksheet, rowNumber)) {
      rows.push(rowNumber);
    }
  }

  return rows;
}

function shiftFormulaRows(formula, fromRow, toRow) {
  const rowDelta = toRow - fromRow;

  return formula.replace(
    CELL_REFERENCE_REGEX,
    (match, absoluteColumn, column, absoluteRow, rowText) => {
      if (absoluteRow === '$') {
        return match;
      }

      const nextRow = Number(rowText) + rowDelta;
      return `${absoluteColumn}${column}${absoluteRow}${nextRow}`;
    }
  );
}

function insertStyledRows(worksheet, insertAtRow, count, templateRowNumber) {
  if (count <= 0) {
    return;
  }

  const templateRow = worksheet.getRow(templateRowNumber);
  const styledColumnCount = Math.max(worksheet.columnCount, 10);

  worksheet.spliceRows(insertAtRow, 0, ...Array.from({ length: count }, () => []));

  for (let offset = 0; offset < count; offset += 1) {
    const newRowNumber = insertAtRow + offset;
    const newRow = worksheet.getRow(newRowNumber);

    if (templateRow.height) {
      newRow.height = templateRow.height;
    }

    for (let column = 1; column <= styledColumnCount; column += 1) {
      const sourceCell = templateRow.getCell(column);
      const targetCell = newRow.getCell(column);
      targetCell.value = null;
      copyCellStyle(sourceCell, targetCell);
    }

    newRow.commit();
  }
}

function clearSharedFormulaMetadata(cell) {
  if (!cell?.model) {
    return;
  }

  delete cell.model.sharedFormula;
  delete cell.model.shareType;
  delete cell.model.ref;
  delete cell.model.si;
}

function applyFormulaPattern(formulaTemplate, rowNumber) {
  return formulaTemplate.replaceAll('{row}', String(rowNumber));
}

function rehydrateSectionFormulas(worksheet, sectionKey, writableRows) {
  if (writableRows.length === 0) {
    return;
  }

  const formulaPattern = SECTION_FORMULA_PATTERNS.get(sectionKey);

  if (!formulaPattern) {
    return;
  }

  for (const rowNumber of writableRows) {
    const row = worksheet.getRow(rowNumber);

    for (const { column, formula } of formulaPattern) {
      const cell = row.getCell(column);
      const result = cell.model ? cell.model.result : undefined;
      cell.value = {
        formula: applyFormulaPattern(formula, rowNumber),
        result
      };
      clearSharedFormulaMetadata(cell);
    }

    row.commit();
  }
}

function rehydrateAllSectionFormulas(worksheet) {
  for (const { key, headerText } of SECTION_DEFINITIONS) {
    const headerRow = findHeaderRow(worksheet, headerText);

    if (!headerRow) {
      continue;
    }

    const sectionStart = headerRow + 1;
    const sectionEnd = findSectionEnd(worksheet, headerRow);
    const writableRows = getWritableRows(worksheet, sectionStart, sectionEnd);
    rehydrateSectionFormulas(worksheet, key, writableRows);
  }
}

function refreshMainSummaryFormulas(worksheet) {
  const subtotalRow = findHeaderRow(worksheet, 'Sub Totals');

  if (!subtotalRow) {
    return;
  }

  const lastDataRow = subtotalRow - 1;
  const differenceRow = subtotalRow + 1;
  const halfDifferenceRow = subtotalRow + 2;
  const percentageRow = subtotalRow + 3;
  const equalizationRow = subtotalRow + 4;
  const totalRow = subtotalRow + 5;
  const finalPercentageRow = subtotalRow + 6;

  worksheet.getRow(subtotalRow).getCell(1).value = {
    formula: `SUM(A3:A${lastDataRow})`
  };
  worksheet.getRow(subtotalRow).getCell(2).value = {
    formula: `SUM(B3:B${lastDataRow})`
  };
  worksheet.getRow(subtotalRow).getCell(3).value = {
    formula: `A${subtotalRow}+B${subtotalRow}`
  };
  worksheet.getRow(subtotalRow).getCell(5).value = {
    formula: `SUM(E3:E${lastDataRow})`
  };
  worksheet.getRow(subtotalRow).getCell(6).value = {
    formula: `SUM(F3:F${lastDataRow})`
  };

  worksheet.getRow(differenceRow).getCell(5).value = {
    formula: `E${subtotalRow}-F${subtotalRow}`
  };

  worksheet.getRow(halfDifferenceRow).getCell(5).value = {
    formula: `E${differenceRow}/2`
  };

  worksheet.getRow(percentageRow).getCell(5).value = {
    formula: `IF(ISERROR(E${subtotalRow}/C${subtotalRow}),0,E${subtotalRow}/C${subtotalRow})`
  };
  worksheet.getRow(percentageRow).getCell(6).value = {
    formula: `IF(ISERROR(F${subtotalRow}/C${subtotalRow}),0,F${subtotalRow}/C${subtotalRow})`
  };

  worksheet.getRow(equalizationRow).getCell(6).value = {
    formula: `E${equalizationRow}*-1`
  };

  worksheet.getRow(totalRow).getCell(5).value = {
    formula: `E${subtotalRow}+E${equalizationRow}`
  };
  worksheet.getRow(totalRow).getCell(6).value = {
    formula: `F${subtotalRow}+F${equalizationRow}`
  };

  worksheet.getRow(finalPercentageRow).getCell(5).value = {
    formula: `IF(ISERROR(E${totalRow}/C${subtotalRow}),0,E${totalRow}/C${subtotalRow})`
  };
  worksheet.getRow(finalPercentageRow).getCell(6).value = {
    formula: `IF(ISERROR(F${totalRow}/C${subtotalRow}),0,F${totalRow}/C${subtotalRow})`
  };
}

function refreshSeparatePropertyTotals(worksheet, sectionHeader) {
  const headerRow = findHeaderRow(worksheet, sectionHeader);

  if (!headerRow) {
    return;
  }

  const sectionStart = headerRow + 1;
  const sectionEnd = findSectionEnd(worksheet, headerRow);
  const totalRow = sectionEnd + 1;

  worksheet.getRow(totalRow).getCell(3).value = {
    formula: `SUM(C${sectionStart}:C${sectionEnd})`
  };
}

function coerceNumber(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value === 'string') {
    const normalizedValue = value.replace(/[$,\s]/g, '');
    if (!normalizedValue) {
      return null;
    }

    const parsed = Number(normalizedValue);
    return Number.isFinite(parsed) ? parsed : null;
  }

  return null;
}

function trimOrNull(value) {
  if (value === null || value === undefined) {
    return null;
  }

  const nextValue = String(value).trim();
  return nextValue ? nextValue : null;
}

function isZeroLike(value) {
  return value === null || value === undefined || value === 0;
}

function normalizeInputRow(rowData) {
  if (!rowData || typeof rowData !== 'object') {
    return null;
  }

  return {
    ...rowData,
    description: trimOrNull(rowData.description),
    notes: trimOrNull(rowData.notes),
    category: trimOrNull(rowData.category),
    market_value: coerceNumber(rowData.market_value),
    debt_balance: coerceNumber(rowData.debt_balance)
  };
}

function isHeadingLikeRow(rowData, targetSectionKey) {
  const description = normalizeHeadingText(rowData.description || '');

  if (!description) {
    return !rowData.notes && isZeroLike(rowData.market_value) && isZeroLike(rowData.debt_balance);
  }

  return (
    KNOWN_HEADING_TEXTS.has(description) ||
    description === normalizeHeadingText(targetSectionKey) ||
    description === normalizeHeadingText(getSectionHeader(targetSectionKey))
  );
}

function getRowSearchText(rowData) {
  return normalizeHeadingText(
    [rowData.description, rowData.notes, rowData.category].filter(Boolean).join(' ')
  );
}

function isCreditCardLikeText(text) {
  return (
    /\bcredit card\b/.test(text) ||
    /\b(american express|amex|visa|mastercard|master card|discover|capital one)\b/.test(text) ||
    (/\b(chase|wells fargo)\b/.test(text) && /\bcard\b/.test(text))
  );
}

function isFederalIncomeTaxLikeText(text) {
  return /\b(irs|federal income tax|federal tax|tax liability|income taxes?)\b/.test(text);
}

function isAttorneyFeeLikeText(text) {
  return /\b(attorney fee|attorney fees|attorney s fee|attorney s fees|expert fee|expert fees|legal fee)\b/.test(text);
}

function isBusinessInterestLikeText(text) {
  return /\b(closely held business|business interest|member interest|ownership interest|partnership interest)\b/.test(text);
}

function isStockOptionLikeText(text) {
  return /\b(stock option|restricted stock)\b/.test(text);
}

function isBonusLikeText(text) {
  return /\bbonus(es)?\b/.test(text);
}

function determineTargetSectionKey(currentSectionKey, rowData) {
  const text = getRowSearchText(rowData);

  for (const rule of ROUTING_RULES) {
    if (rule.match(text)) {
      return rule.key;
    }
  }

  return currentSectionKey;
}

function normalizeDebtValues(rowData, targetSectionKey) {
  if (!DEBT_SECTIONS.has(targetSectionKey) && !isCreditCardLikeText(getRowSearchText(rowData))) {
    return rowData;
  }

  let marketValue = rowData.market_value;
  let debtBalance = rowData.debt_balance;

  if ((debtBalance === null || debtBalance === 0) && marketValue !== null && marketValue !== 0) {
    debtBalance = -Math.abs(marketValue);
    marketValue = 0;
  } else if (debtBalance !== null && debtBalance !== 0) {
    debtBalance = -Math.abs(debtBalance);
  }

  return {
    ...rowData,
    market_value: marketValue,
    debt_balance: debtBalance
  };
}

function normalizeSectionsPayload(rawSections) {
  const orderedSections = new Map(SECTION_DEFINITIONS.map(({ key }) => [key, []]));
  const errors = [];

  for (const [rawSectionKey, rowsData] of Object.entries(rawSections)) {
    const canonicalSectionKey = resolveSectionKey(rawSectionKey);

    if (!canonicalSectionKey) {
      errors.push(
        `Unknown section key: "${rawSectionKey}". Valid keys: ${SECTION_DEFINITIONS.map(({ key }) => key).join(', ')}`
      );
      continue;
    }

    if (!Array.isArray(rowsData)) {
      errors.push(`Section "${rawSectionKey}" must be an array of row objects`);
      continue;
    }

    for (const rawRow of rowsData) {
      const normalizedRow = normalizeInputRow(rawRow);

      if (!normalizedRow) {
        continue;
      }

      const targetSectionKey = determineTargetSectionKey(canonicalSectionKey, normalizedRow);

      if (isHeadingLikeRow(normalizedRow, targetSectionKey)) {
        continue;
      }

      orderedSections.get(targetSectionKey).push(normalizeDebtValues(normalizedRow, targetSectionKey));
    }
  }

  return { orderedSections, errors };
}

function processSection(worksheet, headerText, rows, writableColumns = [1, 2, 4, 7]) {
  const headerRow = findHeaderRow(worksheet, headerText);

  if (!headerRow) {
    throw new Error(`Section header not found: "${headerText}"`);
  }

  const sectionStart = headerRow + 1;
  let sectionEnd = findSectionEnd(worksheet, headerRow);
  let writableRows = getWritableRows(worksheet, sectionStart, sectionEnd);

  if (rows.length > writableRows.length && writableRows.length > 0) {
    const rowsToInsert = rows.length - writableRows.length;
    const templateRowNumber = writableRows[0];
    const insertPosition = writableRows[writableRows.length - 1] + 1;

    insertStyledRows(worksheet, insertPosition, rowsToInsert, templateRowNumber);
    sectionEnd = findSectionEnd(worksheet, headerRow);
    writableRows = getWritableRows(worksheet, sectionStart, sectionEnd);
  }

  if (writableRows.length === 0) {
    throw new Error(`Section "${headerText}" has no writable rows`);
  }

  rehydrateSectionFormulas(worksheet, resolveSectionKey(headerText) || headerText, writableRows);

  for (const rowNumber of writableRows) {
    const row = worksheet.getRow(rowNumber);

    for (const column of writableColumns) {
      clearCellValue(row.getCell(column));
    }
  }

  const rowsToWrite = Math.min(rows.length, writableRows.length);

  for (let index = 0; index < rowsToWrite; index += 1) {
    const rowData = rows[index];
    const row = worksheet.getRow(writableRows[index]);

    if (rowData.market_value !== null) {
      row.getCell(1).value = rowData.market_value;
    }

    if (rowData.debt_balance !== null) {
      row.getCell(2).value = rowData.debt_balance;
    }

    if (rowData.description) {
      row.getCell(4).value = rowData.description;
    }

    if (rowData.notes) {
      row.getCell(7).value = rowData.notes;
    }

    row.commit();
  }
}

export async function buildMediationWorkbook(body) {
  const sections = extractSections(body);

  if (!sections || typeof sections !== 'object') {
    const error = new Error(
      'Request must contain a "sections" object. Accepted formats: { "sections": {...} }, [{ "sections": {...} }], or { "json": { "sections": {...} } }'
    );
    error.statusCode = 400;
    throw error;
  }

  const { orderedSections, errors } = normalizeSectionsPayload(sections);

  if (errors.length > 0) {
    const error = new Error(errors.join(' | '));
    error.statusCode = 400;
    throw error;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.calcProperties.forceFullCalc = true;

  await workbook.xlsx.readFile(TEMPLATE_PATH);

  const worksheet = workbook.getWorksheet(SHEET_NAME);

  if (!worksheet) {
    const error = new Error(`Worksheet "${SHEET_NAME}" not found in template`);
    error.statusCode = 500;
    throw error;
  }

  for (const [sectionKey, rows] of orderedSections.entries()) {
    if (rows.length === 0) {
      continue;
    }

    processSection(worksheet, SECTION_HEADER_BY_KEY.get(sectionKey), rows);
  }

  rehydrateAllSectionFormulas(worksheet);
  refreshMainSummaryFormulas(worksheet);
  refreshSeparatePropertyTotals(worksheet, "Spouse 1's Separate Property");
  refreshSeparatePropertyTotals(worksheet, "Spouse 2's Separate Property");
  refreshSeparatePropertyTotals(worksheet, "Children's Property");

  return workbook;
}

export async function buildMediationWorkbookBuffer(body) {
  const workbook = await buildMediationWorkbook(body);
  return workbook.xlsx.writeBuffer();
}

export function createApp() {
  const app = express();
  app.use(express.json({ limit: '10mb' }));

  app.get('/health', (req, res) => {
    res.json({ ok: true });
  });

  app.post('/generate-mediation-xlsx', async (req, res) => {
    try {
      const buffer = await buildMediationWorkbookBuffer(req.body);
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const filename = `MediationSpreadsheet-${timestamp}.xlsx`;

      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.setHeader('Content-Length', buffer.length);
      res.send(buffer);
    } catch (error) {
      res.status(error.statusCode || 500).json({
        error: error.statusCode === 400 ? 'Invalid request body' : 'Internal server error',
        message: error.message
      });
    }
  });

  app.use((req, res) => {
    res.status(404).json({
      error: 'Not found',
      message: `Endpoint ${req.method} ${req.path} not found`
    });
  });

  app.use((error, req, res, next) => {
    res.status(500).json({
      error: 'Internal server error',
      message: error.message
    });
  });

  return app;
}

export function startServer(port = DEFAULT_PORT) {
  const app = createApp();

  return app.listen(port, () => {
    console.log(`Spreadsheet Builder API running on port ${port}`);
    console.log(`Health check: http://localhost:${port}/health`);
    console.log(`Generate XLSX: POST http://localhost:${port}/generate-mediation-xlsx`);
  });
}

if (process.argv[1] && path.resolve(process.argv[1]) === __filename) {
  startServer();
}
