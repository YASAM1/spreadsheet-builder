import express from 'express';
import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json({ limit: '10mb' }));

const PORT = process.env.PORT || 3000;

// Template path
const TEMPLATE_PATH = path.join(__dirname, 'template', '(API) 2025.07.30.Inventory-MediationSpreadsheetMaster.xlsx');

// Target sheet name
const SHEET_NAME = 'Prop Distr - KB Preferred - Use';

// Section key -> template header text mapping (sections we write data to)
const SECTION_MAPPING = {
  'Real Property': 'Real Property',
  'Bank Accounts': 'Bank Accounts',
  'Retirement Benefits': 'Retirement Benefits',
  'Personal Property': 'Personal Property (Furniture, Electronics, Firearms, Jewelry, etc)',
  'Life Insurance / Annuities': 'Life Insurance / Annuities',
  'Vehicles': 'Vehicles',
  'Credit Cards': 'Credit Cards',
  'Other Debts & Liabilities': 'Other Debts & Liabilities',
  "Spouse 1's Separate Property": "Spouse 1's Separate Property",
  "Spouse 2's Separate Property": "Spouse 2's Separate Property",
  "Children's Property": "Children's Property"
};

// ALL protected headers in the template (must never be overwritten)
const PROTECTED_HEADERS = [
  'Real Property',
  'Bank Accounts',
  'Business Interests',
  'Retirement Benefits',
  'Stock Options',
  'Personal Property (Furniture, Electronics, Firearms, Jewelry, etc)',
  'Life Insurance / Annuities',
  'Vehicles',
  'Credit Cards',
  'Bonuses',
  'Other Debts & Liabilities',
  'Federal Income Taxes',
  "Attorney's & Expert Fees",
  'Sub Totals',
  '$ Difference between Spouse 1 and Spouse 2',
  '$ Difference between Spouse 1 and Spouse 2 ÷ 2',
  '% Distribution before Equalization',
  'Amt needed to Equalize (Use figure from Cell "$ Difference between Spouse 2 and Spouse 2 ÷ 2" and reverse +/-)',
  'Total',
  "Spouse 1's Separate Property",
  "Spouse 2's Separate Property",
  "Children's Property"
];

// All known headers (for detecting section boundaries) - normalized for comparison
const ALL_HEADERS_RAW = PROTECTED_HEADERS.map(h => h.toLowerCase().trim());

/**
 * Normalize apostrophes: convert curly quotes to straight quotes
 */
function normalizeApostrophes(str) {
  if (!str || typeof str !== 'string') return str || '';
  return str.replace(/[\u2018\u2019\u201A\u201B]/g, "'");
}

/**
 * Get section header from mapping, handling curly/straight apostrophe variants
 */
function getSectionHeader(sectionKey) {
  // Try direct match first
  if (SECTION_MAPPING[sectionKey]) {
    return SECTION_MAPPING[sectionKey];
  }
  // Try normalized match
  const normalizedKey = normalizeApostrophes(sectionKey);
  if (SECTION_MAPPING[normalizedKey]) {
    return SECTION_MAPPING[normalizedKey];
  }
  // Try matching against normalized mapping keys
  for (const [key, value] of Object.entries(SECTION_MAPPING)) {
    if (normalizeApostrophes(key) === normalizedKey) {
      return value;
    }
  }
  return null;
}

/**
 * Normalize incoming request body to extract sections object
 */
function extractSections(body) {
  // Case B: Array format [ { "sections": {...} } ]
  if (Array.isArray(body)) {
    if (body.length > 0 && body[0].sections) {
      return body[0].sections;
    }
    if (body.length > 0 && body[0].json && body[0].json.sections) {
      return body[0].json.sections;
    }
    return null;
  }

  // Case A: Direct object { "sections": {...} }
  if (body.sections) {
    return body.sections;
  }

  // Case C: Nested { "json": { "sections": {...} } }
  if (body.json && body.json.sections) {
    return body.json.sections;
  }

  return null;
}

/**
 * Find a header row by scanning column D for text (case-insensitive, trimmed, apostrophe-normalized)
 */
function findHeaderRow(worksheet, headerText) {
  const targetText = normalizeApostrophes(headerText.toLowerCase().trim());
  let foundRow = null;

  worksheet.eachRow((row, rowNumber) => {
    const cellD = row.getCell(4); // Column D
    const cellValue = getCellStringValue(cellD);
    const normalizedCellValue = normalizeApostrophes(cellValue.toLowerCase().trim());
    if (normalizedCellValue === targetText) {
      foundRow = rowNumber;
    }
  });

  return foundRow;
}

/**
 * Get string value from a cell (handling various cell value types)
 */
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
      // Error values (like #REF!, #N/A, etc.)
      if (cell.value.error) {
        return '';
      }
      // Rich text
      if (cell.value.richText) {
        return cell.value.richText.map(rt => rt.text || '').join('');
      }
      // Formula with result
      if (cell.value.formula && cell.value.result !== undefined) {
        if (typeof cell.value.result === 'object' && cell.value.result?.error) {
          return '';
        }
        return String(cell.value.result || '');
      }
      // Hyperlink
      if (cell.value.text) {
        return cell.value.text;
      }
      // Date
      if (cell.value instanceof Date) {
        return cell.value.toISOString();
      }
    }
    return String(cell.value || '');
  } catch (e) {
    return '';
  }
}

/**
 * Check if a value matches any known section header (normalized)
 */
function isKnownHeader(value) {
  if (!value || typeof value !== 'string') return false;
  const normalized = normalizeApostrophes(value.toLowerCase().trim());
  return ALL_HEADERS_RAW.some(h => normalizeApostrophes(h) === normalized);
}

/**
 * Find the end of a section (row before the next known header or end of data)
 */
function findSectionEnd(worksheet, startRow) {
  let endRow = startRow;
  const lastRow = worksheet.lastRow ? worksheet.lastRow.number : startRow;

  for (let rowNum = startRow + 1; rowNum <= lastRow + 1; rowNum++) {
    const row = worksheet.getRow(rowNum);
    const cellD = row.getCell(4);
    const cellValue = getCellStringValue(cellD);

    // Check if this row contains another section header
    if (isKnownHeader(cellValue)) {
      break;
    }
    endRow = rowNum;
  }

  return endRow;
}

/**
 * Copy cell style from source to target
 */
function copyCellStyle(sourceCell, targetCell) {
  if (sourceCell.style) {
    targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
  }
  if (sourceCell.font) {
    targetCell.font = { ...sourceCell.font };
  }
  if (sourceCell.fill) {
    targetCell.fill = { ...sourceCell.fill };
  }
  if (sourceCell.border) {
    targetCell.border = { ...sourceCell.border };
  }
  if (sourceCell.alignment) {
    targetCell.alignment = { ...sourceCell.alignment };
  }
  if (sourceCell.numFmt) {
    targetCell.numFmt = sourceCell.numFmt;
  }
}

/**
 * Copy row styles from source row to target row for specific columns
 */
function copyRowStyles(worksheet, sourceRowNum, targetRowNum, columns) {
  const sourceRow = worksheet.getRow(sourceRowNum);
  const targetRow = worksheet.getRow(targetRowNum);

  columns.forEach(col => {
    const sourceCell = sourceRow.getCell(col);
    const targetCell = targetRow.getCell(col);
    copyCellStyle(sourceCell, targetCell);
  });

  // Copy row height if set
  if (sourceRow.height) {
    targetRow.height = sourceRow.height;
  }
}

/**
 * Clear cell value but keep formatting
 */
function clearCellValue(cell) {
  try {
    // Skip if cell has a formula (preserve formulas)
    if (cell.formula || cell.sharedFormula) {
      return;
    }

    // Simply set value to null - ExcelJS should preserve styling
    cell.value = null;
  } catch (err) {
    // Silently ignore errors when clearing cells with special content
    console.error(`Warning: Could not clear cell: ${err.message}`);
  }
}

/**
 * Check if a row contains a protected header in column D
 */
function isProtectedRow(worksheet, rowNum) {
  const row = worksheet.getRow(rowNum);
  const cellD = row.getCell(4);
  const cellValue = getCellStringValue(cellD);
  return isKnownHeader(cellValue);
}

/**
 * Get list of writable row numbers in a section (excludes protected header rows)
 */
function getWritableRows(worksheet, sectionStart, sectionEnd) {
  const writableRows = [];
  for (let rowNum = sectionStart; rowNum <= sectionEnd; rowNum++) {
    if (!isProtectedRow(worksheet, rowNum)) {
      writableRows.push(rowNum);
    }
  }
  return writableRows;
}

/**
 * Clear any shared formula references that might cause issues
 * This is needed because ExcelJS doesn't handle shared formulas well when rows are inserted
 */
function clearBrokenSharedFormulas(worksheet) {
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      try {
        // If cell has a sharedFormula reference but no actual formula, clear it
        if (cell.sharedFormula && !cell.formula) {
          cell.value = null;
        }
        // If cell has a formula that references a non-existent shared formula master
        if (cell.model && cell.model.sharedFormula !== undefined) {
          // Check if this is a slave formula (has si but no formula text)
          if (cell.model.sharedFormula !== null && !cell.model.formula) {
            cell.value = null;
          }
        }
      } catch (err) {
        // Ignore errors during cleanup
      }
    });
  });
}

/**
 * Insert new rows at a specific position and copy styles from a template row
 * Only copies styles for the data columns, avoiding formula issues
 */
function insertStyledRows(worksheet, insertAtRow, count, templateRowNum) {
  if (count <= 0) return;

  const templateRow = worksheet.getRow(templateRowNum);
  const columnsToStyle = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]; // Style columns A-J

  // Insert empty rows at the specified position
  worksheet.spliceRows(insertAtRow, 0, ...Array(count).fill([]));

  // Copy styles from template row to each new row
  for (let i = 0; i < count; i++) {
    const newRowNum = insertAtRow + i;
    const newRow = worksheet.getRow(newRowNum);

    // Copy row height
    if (templateRow.height) {
      newRow.height = templateRow.height;
    }

    // Copy cell styles (but NOT values or formulas)
    columnsToStyle.forEach(col => {
      try {
        const sourceCell = templateRow.getCell(col);
        const targetCell = newRow.getCell(col);

        // Explicitly clear any formula/sharedFormula that might have been copied
        targetCell.value = null;

        // Copy style properties individually to avoid formula references
        if (sourceCell.font) targetCell.font = { ...sourceCell.font };
        if (sourceCell.fill) targetCell.fill = JSON.parse(JSON.stringify(sourceCell.fill));
        if (sourceCell.border) targetCell.border = JSON.parse(JSON.stringify(sourceCell.border));
        if (sourceCell.alignment) targetCell.alignment = { ...sourceCell.alignment };
        if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
      } catch (err) {
        // Silently continue if style copy fails for a cell
      }
    });

    newRow.commit();
  }

  console.log(`Inserted ${count} new rows at position ${insertAtRow}`);
}

/**
 * Process a single section: clear old data and write new data
 */
function processSection(worksheet, headerText, rows, allColumns = [1, 2, 4, 7]) {
  const headerRow = findHeaderRow(worksheet, headerText);

  if (!headerRow) {
    throw new Error(`Section header not found: "${headerText}"`);
  }

  const sectionStart = headerRow + 1;
  let sectionEnd = findSectionEnd(worksheet, headerRow);

  // Get only the rows that are safe to write to (not protected headers)
  let writableRows = getWritableRows(worksheet, sectionStart, sectionEnd);
  let existingRowCount = writableRows.length;
  const neededRowCount = rows.length;

  // If we need more rows, insert them at the end of the current writable section
  if (neededRowCount > existingRowCount && existingRowCount > 0) {
    const rowsToInsert = neededRowCount - existingRowCount;
    const templateRowNum = writableRows[0]; // Use first writable row as style template
    const insertPosition = writableRows[writableRows.length - 1] + 1; // Insert after last writable row

    try {
      insertStyledRows(worksheet, insertPosition, rowsToInsert, templateRowNum);

      // Update section end and writable rows after insertion
      sectionEnd = findSectionEnd(worksheet, headerRow);
      writableRows = getWritableRows(worksheet, sectionStart, sectionEnd);
      existingRowCount = writableRows.length;

      console.log(`Section "${headerText}": Added ${rowsToInsert} rows. Now has ${existingRowCount} writable rows.`);
    } catch (err) {
      console.error(`Error inserting rows for section "${headerText}": ${err.message}`);
      console.warn(`Proceeding with ${existingRowCount} available rows.`);
    }
  }

  // Clear existing data rows (columns A, B, D, G only) - only writable rows
  for (const rowNum of writableRows) {
    const row = worksheet.getRow(rowNum);
    allColumns.forEach(col => {
      try {
        const cell = row.getCell(col);
        // Only clear if it's not a formula (preserve formulas in other scenarios)
        if (!cell.formula) {
          clearCellValue(cell);
        }
      } catch (err) {
        console.error(`Error clearing cell at row ${rowNum}, col ${col}: ${err.message}`);
      }
    });
  }

  // Final check if we still don't have enough rows
  if (neededRowCount > existingRowCount) {
    console.warn(`Warning: Section "${headerText}" has ${existingRowCount} available rows but needs ${neededRowCount}. Extra data will be truncated.`);
  }

  // Write new data to writable rows only (limited to available rows)
  const rowsToWrite = Math.min(rows.length, existingRowCount);

  for (let index = 0; index < rowsToWrite; index++) {
    const rowData = rows[index];
    const targetRowNum = writableRows[index];
    const row = worksheet.getRow(targetRowNum);

    try {
      // Column A = market_value
      if (rowData.market_value !== undefined && rowData.market_value !== null) {
        row.getCell(1).value = rowData.market_value;
      }

      // Column B = debt_balance
      if (rowData.debt_balance !== undefined && rowData.debt_balance !== null) {
        row.getCell(2).value = rowData.debt_balance;
      }

      // Column D = description
      if (rowData.description !== undefined && rowData.description !== null) {
        row.getCell(4).value = rowData.description;
      }

      // Column G = notes
      if (rowData.notes !== undefined && rowData.notes !== null) {
        row.getCell(7).value = rowData.notes;
      }

      row.commit();
    } catch (err) {
      console.error(`Error writing to row ${targetRowNum}: ${err.message}`);
    }
  }
}

// =====================
// ROUTES
// =====================

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ ok: true });
});

// Main endpoint to generate mediation XLSX
app.post('/generate-mediation-xlsx', async (req, res) => {
  try {
    // Extract and validate sections
    const sections = extractSections(req.body);

    if (!sections || typeof sections !== 'object') {
      return res.status(400).json({
        error: 'Invalid request body',
        message: 'Request must contain a "sections" object. Accepted formats: { "sections": {...} }, [{ "sections": {...} }], or { "json": { "sections": {...} } }'
      });
    }

    // Load the template workbook
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(TEMPLATE_PATH);

    // Get the target worksheet
    const worksheet = workbook.getWorksheet(SHEET_NAME);

    if (!worksheet) {
      return res.status(500).json({
        error: 'Template error',
        message: `Worksheet "${SHEET_NAME}" not found in template`
      });
    }

    // Process each section in the input
    const errors = [];

    for (const [sectionKey, rowsData] of Object.entries(sections)) {
      // Get the template header text for this section key (handles curly apostrophes)
      const headerText = getSectionHeader(sectionKey);

      if (!headerText) {
        errors.push(`Unknown section key: "${sectionKey}". Valid keys: ${Object.keys(SECTION_MAPPING).join(', ')}`);
        continue;
      }

      if (!Array.isArray(rowsData)) {
        errors.push(`Section "${sectionKey}" must be an array of row objects`);
        continue;
      }

      // Skip empty sections
      if (rowsData.length === 0) {
        continue;
      }

      try {
        processSection(worksheet, headerText, rowsData);
      } catch (err) {
        errors.push(`Error processing section "${sectionKey}": ${err.message}`);
      }
    }

    // If there were errors, return 500 with details
    if (errors.length > 0) {
      return res.status(500).json({
        error: 'Processing errors',
        messages: errors
      });
    }

    // Clean up any broken shared formula references before generating output
    clearBrokenSharedFormulas(worksheet);

    // Generate the output buffer
    const buffer = await workbook.xlsx.writeBuffer();

    // Set response headers for file download
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const filename = `MediationSpreadsheet-${timestamp}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', buffer.length);

    res.send(buffer);

  } catch (err) {
    console.error('Error generating spreadsheet:', err);
    res.status(500).json({
      error: 'Internal server error',
      message: err.message
    });
  }
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({
    error: 'Not found',
    message: `Endpoint ${req.method} ${req.path} not found`
  });
});

// Error handler
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({
    error: 'Internal server error',
    message: err.message
  });
});

// Start server
app.listen(PORT, () => {
  console.log(`Spreadsheet Builder API running on port ${PORT}`);
  console.log(`Health check: http://localhost:${PORT}/health`);
  console.log(`Generate XLSX: POST http://localhost:${PORT}/generate-mediation-xlsx`);
});
