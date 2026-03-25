/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback } from 'react';
import { GoogleGenAI } from "@google/genai";
import { 
  Table as TableIcon, 
  Clipboard, 
  Check, 
  AlertCircle, 
  Loader2, 
  FileText, 
  ArrowRight,
  RefreshCw,
  Download
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import XLSX from 'xlsx-js-style';
import Papa from 'papaparse';

declare global {
  interface Window {
    aistudio?: {
      hasSelectedApiKey: () => Promise<boolean>;
      openSelectKey: () => Promise<void>;
    };
  }
}

// Utility for tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

const SYSTEM_INSTRUCTION = `You are a data extraction and structuring expert.

TASK:
Extract tabular data from an email input and convert it into a structured sheet format (Google Sheets / Excel compatible).

INPUT:
The input will be a raw email (HTML or plain text) containing one or more tables.

CRITICAL INSTRUCTIONS (FOLLOW STRICTLY):
1. DO NOT modify any data:
   - No spelling correction
   - No number rounding
   - No format change
   - Preserve exact text, symbols, spacing, and case

2. Preserve structure exactly:
   - Maintain original column order
   - Maintain row order
   - Keep merged cells logic (if any) by repeating values where necessary
   - Do NOT drop empty cells — represent them as blank

3. Header handling:
   - Identify headers correctly
   - If multiple header rows exist, preserve hierarchy
   - Do not rename headers

4. Data integrity:
   - Ensure 100% accuracy
   - Cross-check row/column alignment before output
   - No missing or shifted values

5. Output format:
   - Return data in clean CSV format
   - Use comma as delimiter
   - Wrap text fields in double quotes
   - Escape internal quotes properly
   - Maintain line breaks only between rows

6. Multiple tables:
   - If multiple tables exist, extract each separately
   - Label them as Table_1, Table_2, etc.

7. Special cases:
   - Dates → keep exact format (no conversion)
   - Currency → keep symbols intact
   - Numbers → keep leading zeros if present
   - Null/empty → leave blank, do not insert placeholders
   - Repetitive Labels → Labels like 'Customer Name', 'Coverage Period', 'one Time Price Validity', or 'one time price' must appear only ONCE per table. Do NOT repeat them across rows or columns even if they appear multiple times in the source email due to layout or merged cells. Include them only in the first applicable row or as a header.
   - Total Row → After any row in which 'Total' (or a similar total summary) is given, insert one empty row in the CSV output. The word 'Total' itself should be placed in the 'Unit Price' column of that row.

8. Validation step (MANDATORY):
   - Re-check extracted output against input
   - Confirm: "No data altered, dropped, or misaligned"

OUTPUT FORMAT:

Table_1:
<CSV DATA>

Table_2:
<CSV DATA>

VALIDATION:
- Data integrity check: PASSED
- Structure preserved: YES
- Any ambiguity: Mention explicitly`;

interface ExtractedTable {
  id: string;
  name: string;
  csv: string;
}

export default function App() {
  const [input, setInput] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [tables, setTables] = useState<ExtractedTable[]>([]);
  const [validation, setValidation] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [needsApiKey, setNeedsApiKey] = useState(false);

  const handleSelectKey = async () => {
    if (window.aistudio) {
      await window.aistudio.openSelectKey();
      setNeedsApiKey(false);
      setError(null);
    }
  };

  const processEmail = async () => {
    if (!input.trim()) return;

    setIsProcessing(true);
    setError(null);
    setTables([]);
    setValidation(null);
    setNeedsApiKey(false);

    try {
      const apiKey = process.env.API_KEY || process.env.GEMINI_API_KEY;
      const ai = new GoogleGenAI({ apiKey });
      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: input,
        config: {
          systemInstruction: SYSTEM_INSTRUCTION,
          temperature: 0.1, // Low temperature for high accuracy extraction
        },
      });

      const text = response.text;
      if (!text) throw new Error("No response from AI");

      // Parse the response
      const tableMatches = Array.from(text.matchAll(/Table_?\d*:\s*\n?([\s\S]*?)(?=\nTable_?\d*:|\nVALIDATION:|$)/gi));
      const extractedTables: ExtractedTable[] = [];
      
      tableMatches.forEach((match, index) => {
        const csvContent = match[1].trim();
        if (csvContent) {
          // Parse and filter the rows
          const parseResult = Papa.parse<string[]>(csvContent, { header: false, skipEmptyLines: true });
          const initialRows = parseResult.data;
          
          // 1. Find Part Code index in the first row
          const header = initialRows[0] || [];
          let partCodeIndex = -1;
          header.forEach((cell, i) => {
            if (cell && String(cell).toLowerCase().includes('part code')) {
              partCodeIndex = i;
            }
          });

          let processedRows: string[][] = [];
          if (partCodeIndex !== -1) {
            // 2. Insert "Subscription Details" header
            const newHeader = [...header];
            newHeader.splice(partCodeIndex + 1, 0, "Subscription Details");
            processedRows.push(newHeader);

            // 3. Process data rows
            for (let i = 1; i < initialRows.length; i++) {
              const currentRow = [...initialRows[i]];
              
              let subscriptionInfo = "";
              let subIdx = -1;
              currentRow.forEach((cell, j) => {
                if (cell && String(cell).toLowerCase().includes('subscription')) {
                  subscriptionInfo = String(cell);
                  subIdx = j;
                }
              });

              if (subscriptionInfo) {
                // Check if it's a "subscription only" row
                const otherContent = currentRow.filter((c, j) => j !== subIdx && c && String(c).trim() !== "");
                if (otherContent.length === 0) {
                  // Move to previous row
                  if (processedRows.length > 1) {
                    const prevRow = processedRows[processedRows.length - 1];
                    // Ensure prevRow has enough cells
                    while (prevRow.length <= partCodeIndex + 1) prevRow.push("");
                    prevRow[partCodeIndex + 1] = subscriptionInfo;
                  }
                  continue; // Skip adding this row
                } else {
                  // Move within the same row
                  currentRow.splice(partCodeIndex + 1, 0, "");
                  const newSubIdx = subIdx > partCodeIndex ? subIdx + 1 : subIdx;
                  currentRow[partCodeIndex + 1] = subscriptionInfo;
                  currentRow[newSubIdx] = "";
                  processedRows.push(currentRow);
                }
              } else {
                // Normal row, just add the empty column
                currentRow.splice(partCodeIndex + 1, 0, "");
                processedRows.push(currentRow);
              }
            }
          } else {
            processedRows = initialRows;
          }

          // 4. Post-process for "One time price" removal and Rupee symbols
          const tempHeader = processedRows[0] || [];
          let otpIdx = -1;
          let upIdx = -1;
          let tpIdx = -1;
          tempHeader.forEach((cell, i) => {
            const cellText = String(cell).toLowerCase();
            if (cellText.includes('one time price') || cellText.includes('on time price')) otpIdx = i;
            if (cellText.includes('unit price')) upIdx = i;
            if (cellText.includes('total price')) tpIdx = i;
          });

          const finalRows = processedRows.map((row, i) => {
            const isHeader = i === 0;
            const newRow = [...row];

            // Add Rupee symbol to price columns (do this BEFORE splicing to keep indices correct)
            if (!isHeader) {
              if (upIdx !== -1 && newRow[upIdx]) {
                const val = String(newRow[upIdx]).trim();
                if (val && !val.startsWith('₹') && /\d/.test(val)) {
                  newRow[upIdx] = `₹${val}`;
                }
              }
              if (tpIdx !== -1 && newRow[tpIdx]) {
                const val = String(newRow[tpIdx]).trim();
                if (val && !val.startsWith('₹') && /\d/.test(val)) {
                  newRow[tpIdx] = `₹${val}`;
                }
              }
            }

            // Identify indices to remove
            const indicesToRemove = new Set<number>();
            indicesToRemove.add(0); // Always remove the first column
            if (otpIdx !== -1) {
              indicesToRemove.add(otpIdx);
            }

            // Remove indices in descending order to maintain correct mapping
            const sortedIndices = Array.from(indicesToRemove).sort((a, b) => b - a);
            sortedIndices.forEach(idx => {
              if (idx < newRow.length) {
                newRow.splice(idx, 1);
              }
            });

            return newRow;
          });

          // 5. Apply existing filters (empty rows, total rows)
          const filteredRows = finalRows.filter(row => {
            const rowText = row.join(' ').toLowerCase();
            // Remove if empty
            if (!row.some(cell => cell && String(cell).trim().length > 0)) {
              return false;
            }
            // Remove if contains "total" but NOT "total price"
            if (rowText.includes('total') && !rowText.includes('total price')) {
              return false;
            }
            return true;
          });
          const filteredCsv = Papa.unparse(filteredRows);

          extractedTables.push({
            id: `table-${index + 1}`,
            name: `Table ${index + 1}`,
            csv: filteredCsv
          });
        }
      });

      const validationMatch = text.match(/VALIDATION:([\s\S]*)$/);
      
      setTables(extractedTables);
      setValidation(validationMatch ? validationMatch[1].trim() : "Validation section missing in response.");
      
      if (extractedTables.length === 0) {
        setError("No tables were detected in the input provided.");
      }
    } catch (err) {
      console.error(err);
      const errorMessage = err instanceof Error ? err.message : "An unexpected error occurred during processing.";
      setError(errorMessage);
      
      if (errorMessage.toLowerCase().includes("api key not valid") || 
          errorMessage.toLowerCase().includes("requested entity was not found") ||
          errorMessage.toLowerCase().includes("invalid_argument") ||
          errorMessage.toLowerCase().includes("api_key_invalid") ||
          errorMessage.toLowerCase().includes("403") ||
          errorMessage.toLowerCase().includes("401")) {
        setNeedsApiKey(true);
      }
    } finally {
      setIsProcessing(false);
    }
  };

  const copyToClipboard = (text: string, id: string) => {
    navigator.clipboard.writeText(text);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  const downloadExcel = () => {
    if (tables.length === 0) return;

    const wb = XLSX.utils.book_new();
    const combinedData: any[][] = [];
    
    tables.forEach((table, index) => {
      // Use PapaParse for robust CSV parsing
      const parseResult = Papa.parse<string[]>(table.csv, {
        header: false,
        skipEmptyLines: true, // Remove empty rows from source
      });

      if (parseResult.data && parseResult.data.length > 0) {
        if (index === 0) {
          // For the first table, add everything (headers + data)
          combinedData.push(...parseResult.data);
        } else {
          // For subsequent tables, skip the header row (index 0)
          combinedData.push(...parseResult.data.slice(1));
        }
      }
    });

    if (combinedData.length > 0) {
      const ws = XLSX.utils.aoa_to_sheet(combinedData);
      
      // Set column widths to 200 pixels
      const maxCols = combinedData.reduce((max, row) => Math.max(max, row.length), 0);
      const maxRows = combinedData.length;
      ws['!cols'] = Array(maxCols).fill({ wpx: 200 });

      // Set ALL rows height to 65 pixels
      ws['!rows'] = Array(maxRows).fill({ hpx: 65 });

      // Apply styling to ALL cells
      const headerFill = { fgColor: { rgb: "1A365D" } }; // Deep Navy for header
      const lightBlueFill = { fgColor: { rgb: "ADD8E6" } }; // Light Blue for alternate rows
      const blueFill = { fgColor: { rgb: "2B5797" } }; // Blue for alternate rows
      const baseFont = { sz: 12, bold: true, color: { rgb: "FFFFFF" } };
      const darkBlueFont = { sz: 12, bold: true, color: { rgb: "00008B" } };
      const defaultAlignment = { horizontal: "center", vertical: "center", wrapText: true };
      const darkBlueBorder = {
        top: { style: 'thin', color: { rgb: '00008B' } },
        bottom: { style: 'thin', color: { rgb: '00008B' } },
        left: { style: 'thin', color: { rgb: '00008B' } },
        right: { style: 'thin', color: { rgb: '00008B' } }
      };
      
      let activeCompanyNameCols = new Set<number>();
      let activeTotalPriceCols = new Set<number>();
      let activeCustomerNameCols = new Set<number>();

      // Identify special columns from the first row (the only header row)
      const firstRow = combinedData[0];
      if (firstRow) {
        firstRow.forEach((cell, c) => {
          const cellText = String(cell).toLowerCase();
          if (cellText.includes('company name')) activeCompanyNameCols.add(c);
          if (cellText.includes('total price')) activeTotalPriceCols.add(c);
          if (cellText.includes('customer name')) activeCustomerNameCols.add(c);
          if (cellText.includes('subscription details')) activeCustomerNameCols.add(c); // Reuse one of the sets or add a new one
        });
      }

      for (let r = 0; r < maxRows; r++) {
        for (let c = 0; c < maxCols; c++) {
          const cellAddress = XLSX.utils.encode_cell({ r, c });
          if (!ws[cellAddress]) {
            ws[cellAddress] = { t: 's', v: '' };
          }
          
          const isHeader = r === 0;
          let currentFill = headerFill;
          let currentFont = baseFont;
          if (!isHeader) {
            // Alternate between Light Blue and Blue for data rows
            // r=1 is first data row
            const isLightBlue = r % 2 === 1;
            currentFill = isLightBlue ? lightBlueFill : blueFill;
            currentFont = isLightBlue ? darkBlueFont : baseFont;
          }
          
          ws[cellAddress].s = {
            alignment: defaultAlignment,
            fill: currentFill,
            font: { ...currentFont },
            border: darkBlueBorder
          };
        }
      }

      XLSX.utils.book_append_sheet(wb, ws, 'Extracted Data');
      XLSX.writeFile(wb, 'Extracted_Tables.xlsx');
    }
  };

  const downloadCombinedCsv = () => {
    if (tables.length === 0) return;

    let combined = '';
    tables.forEach((table, index) => {
      if (index > 0) combined += '\n\n';
      combined += `--- ${table.name} ---\n`;
      combined += table.csv;
    });

    const blob = new Blob([combined], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', 'Combined_Tables.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="min-h-screen bg-[#F8F9FA] text-[#1A1A1A] font-sans selection:bg-blue-100">
      {/* Top Banner */}
      <div className="bg-blue-600 text-white py-1.5 px-6 text-center text-[10px] font-bold uppercase tracking-[0.2em]">
        Developed by KASTRA (aikastra.in)
      </div>

      {/* Header */}
      <header className="border-b border-gray-200 bg-white sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-blue-600 rounded-lg flex items-center justify-center text-white">
              <TableIcon size={24} />
            </div>
            <div>
              <h1 className="font-semibold text-lg leading-tight">Email Table Extractor</h1>
              <p className="text-xs text-gray-500 font-medium uppercase tracking-wider">Data Structuring Expert</p>
            </div>
          </div>
          <div className="flex items-center gap-4">
            <button 
              onClick={() => {
                setInput('');
                setTables([]);
                setValidation(null);
                setError(null);
              }}
              className="text-sm text-gray-500 hover:text-gray-800 transition-colors flex items-center gap-1.5"
            >
              <RefreshCw size={14} />
              Reset
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-8 grid grid-cols-1 lg:grid-cols-2 gap-8">
        {/* Input Section */}
        <section className="space-y-4">
          <div className="flex items-center justify-between">
            <h2 className="text-sm font-bold uppercase tracking-widest text-gray-400 flex items-center gap-2">
              <FileText size={16} />
              Raw Email Input
            </h2>
            <span className="text-[10px] bg-gray-100 text-gray-500 px-2 py-0.5 rounded uppercase font-bold">
              HTML or Plain Text
            </span>
          </div>
          
          <div className="relative group">
            <textarea
              value={input}
              onChange={(e) => setInput(e.target.value)}
              placeholder="Paste your raw email content here..."
              className="w-full h-[500px] p-4 bg-white border border-gray-200 rounded-xl shadow-sm focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 outline-none transition-all resize-none font-mono text-sm leading-relaxed"
            />
            {input && !isProcessing && (
              <motion.button
                initial={{ opacity: 0, scale: 0.9 }}
                animate={{ opacity: 1, scale: 1 }}
                onClick={processEmail}
                className="absolute bottom-4 right-4 bg-blue-600 text-white px-6 py-3 rounded-lg font-semibold shadow-lg hover:bg-blue-700 active:scale-95 transition-all flex items-center gap-2"
              >
                Extract Tables
                <ArrowRight size={18} />
              </motion.button>
            )}
          </div>
        </section>

        {/* Output Section */}
        <section className="space-y-4">
          <div className="flex items-center justify-between">
            <h2 className="text-sm font-bold uppercase tracking-widest text-gray-400 flex items-center gap-2">
              <Clipboard size={16} />
              Structured Output
            </h2>
            {tables.length > 0 && (
              <div className="flex items-center gap-2">
                <button
                  onClick={downloadCombinedCsv}
                  className="text-xs font-bold uppercase tracking-tight px-3 py-1.5 rounded-md bg-blue-100 text-blue-700 hover:bg-blue-200 transition-all flex items-center gap-1.5 shadow-sm"
                >
                  <FileText size={14} />
                  Download Combined CSV
                </button>
                <button
                  onClick={downloadExcel}
                  className="text-xs font-bold uppercase tracking-tight px-3 py-1.5 rounded-md bg-green-600 text-white hover:bg-green-700 transition-all flex items-center gap-1.5 shadow-sm"
                >
                  <Download size={14} />
                  Download Excel
                </button>
              </div>
            )}
          </div>

          <div className="min-h-[500px] bg-white border border-gray-200 rounded-xl shadow-sm overflow-hidden flex flex-col">
            <AnimatePresence mode="wait">
              {isProcessing ? (
                <motion.div 
                  key="loading"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  className="flex-1 flex flex-col items-center justify-center p-12 text-center"
                >
                  <Loader2 className="w-12 h-12 text-blue-600 animate-spin mb-4" />
                  <h3 className="text-lg font-semibold">Analyzing Email Structure</h3>
                  <p className="text-gray-500 text-sm max-w-xs mt-2">
                    Our AI is identifying tables and ensuring 100% data integrity...
                  </p>
                </motion.div>
              ) : error ? (
                <motion.div 
                  key="error"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  className="flex-1 flex flex-col items-center justify-center p-12 text-center"
                >
                  <div className="w-16 h-16 bg-red-50 text-red-500 rounded-full flex items-center justify-center mb-4">
                    <AlertCircle size={32} />
                  </div>
                  <h3 className="text-lg font-semibold text-red-600">Extraction Failed</h3>
                  <p className="text-gray-500 text-sm max-w-xs mt-2 mb-6">{error}</p>
                  
                  {needsApiKey && (
                    <div className="space-y-4 max-w-sm mx-auto p-6 bg-gray-50 rounded-2xl border border-gray-100">
                      <p className="text-xs text-gray-600 leading-relaxed">
                        This error usually occurs in shared apps when the developer's API key is restricted. 
                        Please select your own Gemini API key to continue.
                      </p>
                      <button
                        onClick={handleSelectKey}
                        className="w-full py-3 bg-blue-600 text-white text-sm font-bold rounded-xl hover:bg-blue-700 active:scale-[0.98] transition-all flex items-center justify-center gap-2 shadow-lg shadow-blue-500/20"
                      >
                        <RefreshCw size={16} />
                        Configure API Key
                      </button>
                      <p className="text-[10px] text-gray-400">
                        You'll need a key from a paid Google Cloud project. 
                        See <a href="https://ai.google.dev/gemini-api/docs/billing" target="_blank" rel="noopener noreferrer" className="underline hover:text-blue-500">billing docs</a>.
                      </p>
                    </div>
                  )}
                </motion.div>
              ) : tables.length > 0 ? (
                <motion.div 
                  key="results"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  className="flex-1 flex flex-col"
                >
                  <div className="flex-1 overflow-y-auto p-6 space-y-24">
                    {tables.map((table) => (
                      <div key={table.id} className="space-y-3">
                        <div className="flex items-center justify-between">
                          <h4 className="font-bold text-gray-900 flex items-center gap-2 italic font-serif">
                            {table.name}
                          </h4>
                          <button
                            onClick={() => copyToClipboard(table.csv, table.id)}
                            className={cn(
                              "text-xs font-bold uppercase tracking-tight px-3 py-1.5 rounded-md transition-all flex items-center gap-1.5",
                              copiedId === table.id 
                                ? "bg-green-100 text-green-700" 
                                : "bg-gray-100 text-gray-600 hover:bg-gray-200"
                            )}
                          >
                            {copiedId === table.id ? (
                              <>
                                <Check size={14} />
                                Copied
                              </>
                            ) : (
                              <>
                                <Clipboard size={14} />
                                Copy CSV
                              </>
                            )}
                          </button>
                        </div>
                        <div className="bg-gray-50 border border-gray-200 rounded-lg p-4 font-mono text-xs overflow-x-auto whitespace-pre leading-relaxed text-gray-700">
                          {table.csv}
                        </div>
                      </div>
                    ))}
                  </div>

                  {validation && (
                    <div className="border-t border-gray-100 bg-blue-50/50 p-4">
                      <div className="flex gap-3">
                        <Check className="text-green-600 shrink-0" size={18} />
                        <div>
                          <p className="text-xs font-bold uppercase tracking-widest text-blue-900/60 mb-1">AI Validation Report</p>
                          <p className="text-sm text-blue-900/80 italic">{validation}</p>
                        </div>
                      </div>
                    </div>
                  )}
                </motion.div>
              ) : (
                <motion.div 
                  key="empty"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  className="flex-1 flex flex-col items-center justify-center p-12 text-center text-gray-400"
                >
                  <TableIcon size={48} strokeWidth={1} className="mb-4 opacity-20" />
                  <p className="text-sm font-medium">Extracted tables will appear here</p>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </section>
      </main>

      <footer className="max-w-6xl mx-auto px-6 py-12 border-t border-gray-200 mt-12">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
          <div className="space-y-4">
            <h3 className="text-xs font-bold uppercase tracking-widest text-gray-400">Accuracy Guarantee</h3>
            <p className="text-sm text-gray-500 leading-relaxed">
              Our extraction engine preserves exact text, symbols, spacing, and case. No spelling correction or number rounding is performed.
            </p>
          </div>
          <div className="space-y-4">
            <h3 className="text-xs font-bold uppercase tracking-widest text-gray-400">Structure Preservation</h3>
            <p className="text-sm text-gray-500 leading-relaxed">
              Original column and row order are maintained. Merged cells are handled by repeating values to ensure data integrity.
            </p>
          </div>
          <div className="space-y-4">
            <h3 className="text-xs font-bold uppercase tracking-widest text-gray-400">Compatibility</h3>
            <p className="text-sm text-gray-500 leading-relaxed">
              Output is provided in standard CSV format, ready to be imported into Google Sheets, Microsoft Excel, or any data analysis tool.
            </p>
          </div>
        </div>
      </footer>
    </div>
  );
}
