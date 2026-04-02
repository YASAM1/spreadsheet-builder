import assert from 'node:assert/strict';
import test from 'node:test';

import ExcelJS from 'exceljs';

import {
  buildMediationWorkbookBuffer,
  SECTION_DEFINITIONS,
  SHEET_NAME
} from '../src/index.js';

const KNOWN_HEADERS = SECTION_DEFINITIONS.map(({ headerText }) => headerText).concat([
  'Sub Totals',
  'Total'
]);

function normalizeHeadingText(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[\u2018\u2019\u201A\u201B]/g, "'")
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function findRowByHeader(worksheet, headerText) {
  const target = normalizeHeadingText(headerText);

  let found = null;
  worksheet.eachRow((row, rowNumber) => {
    const value = row.getCell(4).text || row.getCell(4).value || '';
    if (normalizeHeadingText(value) === target) {
      found = rowNumber;
    }
  });

  return found;
}

function findSectionRows(worksheet, headerText) {
  const headerRow = findRowByHeader(worksheet, headerText);
  assert.ok(headerRow, `Missing header row for ${headerText}`);

  const rows = [];

  for (
    let rowNumber = headerRow + 1;
    rowNumber <= (worksheet.lastRow ? worksheet.lastRow.number : headerRow);
    rowNumber += 1
  ) {
    const description = worksheet.getRow(rowNumber).getCell(4).text || '';
    if (KNOWN_HEADERS.includes(description.trim())) {
      break;
    }
    rows.push(rowNumber);
  }

  return rows;
}

async function loadWorksheet(payload) {
  const buffer = await buildMediationWorkbookBuffer(payload);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook.getWorksheet(SHEET_NAME);
}

test('rehydrates formulas when sections outgrow the template', async () => {
  const worksheet = await loadWorksheet({
    sections: {
      'Personal Property': Array.from({ length: 10 }, (_, index) => ({
        description: `Sporting Item ${index + 1}`,
        market_value: 1000 + index,
        debt_balance: 0,
        notes: `Test item ${index + 1}`
      }))
    }
  });

  const sectionRows = findSectionRows(
    worksheet,
    'Personal Property (Furniture, Electronics, Firearms, Jewelry, etc)'
  );

  assert.equal(sectionRows.length, 10);

  for (const rowNumber of sectionRows) {
    assert.equal(worksheet.getCell(`C${rowNumber}`).formula, `A${rowNumber}+B${rowNumber}`);
    assert.equal(worksheet.getCell(`E${rowNumber}`).formula, `C${rowNumber}*0.5`);
    assert.equal(worksheet.getCell(`F${rowNumber}`).formula, `C${rowNumber}-E${rowNumber}`);
  }

  const subtotalRow = findRowByHeader(worksheet, 'Sub Totals');
  assert.ok(subtotalRow, 'Expected subtotal row');
  assert.equal(worksheet.getCell(`A${subtotalRow}`).formula, `SUM(A3:A${subtotalRow - 1})`);
  assert.equal(worksheet.getCell(`B${subtotalRow}`).formula, `SUM(B3:B${subtotalRow - 1})`);
  assert.equal(worksheet.getCell(`C${subtotalRow}`).formula, `A${subtotalRow}+B${subtotalRow}`);
});

test('suppresses heading rows and reroutes obvious miscategorized liabilities', async () => {
  const worksheet = await loadWorksheet({
    sections: {
      'Real Property': [
        {
          description: 'Real Properties',
          market_value: 0,
          debt_balance: 0,
          notes: ''
        },
        {
          description: '123 Main Street',
          market_value: 250000,
          debt_balance: -90000,
          notes: 'Mortgage'
        }
      ],
      'Life Insurance / Annuities': [
        {
          description: 'Credit Cards',
          market_value: 0,
          debt_balance: 0,
          notes: ''
        },
        {
          description: 'American Express x1009',
          market_value: 4567,
          debt_balance: 0,
          notes: 'Imported into the wrong section'
        },
        {
          description: 'Closely Held Business Interest - Acme LLC',
          market_value: 80000,
          debt_balance: 0,
          notes: ''
        }
      ],
      'Other Debts & Liabilities': [
        {
          description: 'IRS income tax liability 2024',
          market_value: 1250,
          debt_balance: 0,
          notes: ''
        }
      ]
    }
  });

  const realPropertyRows = findSectionRows(worksheet, 'Real Property');
  const realPropertyDescriptions = realPropertyRows
    .map((rowNumber) => worksheet.getCell(`D${rowNumber}`).text)
    .filter(Boolean);
  assert.deepEqual(realPropertyDescriptions, ['123 Main Street']);

  const creditCardRows = findSectionRows(worksheet, 'Credit Cards');
  const creditCardDescription = creditCardRows
    .map((rowNumber) => ({
      rowNumber,
      description: worksheet.getCell(`D${rowNumber}`).text
    }))
    .find(({ description }) => description === 'American Express x1009');

  assert.ok(creditCardDescription, 'Expected rerouted credit card row');
  assert.equal(worksheet.getCell(`A${creditCardDescription.rowNumber}`).value, 0);
  assert.equal(worksheet.getCell(`B${creditCardDescription.rowNumber}`).value, -4567);

  const businessRows = findSectionRows(worksheet, 'Business Interests');
  const businessDescriptions = businessRows
    .map((rowNumber) => worksheet.getCell(`D${rowNumber}`).text)
    .filter(Boolean);
  assert.ok(businessDescriptions.includes('Closely Held Business Interest - Acme LLC'));

  const taxRows = findSectionRows(worksheet, 'Federal Income Taxes');
  const taxEntry = taxRows
    .map((rowNumber) => ({
      rowNumber,
      description: worksheet.getCell(`D${rowNumber}`).text
    }))
    .find(({ description }) => description === 'IRS income tax liability 2024');

  assert.ok(taxEntry, 'Expected rerouted tax row');
  assert.equal(worksheet.getCell(`A${taxEntry.rowNumber}`).value, 0);
  assert.equal(worksheet.getCell(`B${taxEntry.rowNumber}`).value, -1250);
});
