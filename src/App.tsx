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

  const processEmail = async () => {
    if (!input.trim()) return;

    setIsProcessing(true);
    setError(null);
    setTables([]);
    setValidation(null);

    try {
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
      const response = await ai.models.generateContent({
        model: "gemini-3.1-pro-preview",
        contents: input,
        config: {
          systemInstruction: SYSTEM_INSTRUCTION,
          temperature: 0.1, // Low temperature for high accuracy extraction
        },
      });

      const text = response.text;
      if (!text) throw new Error("No response from AI");

      // Parse the response
      const tableMatches = text.matchAll(/Table_(\d+):\n([\s\S]*?)(?=\nTable_\d+:|\nVALIDATION:|$)/g);
      const extractedTables: ExtractedTable[] = [];
      
      for (const match of tableMatches) {
        extractedTables.push({
          id: `table-${match[1]}`,
          name: `Table ${match[1]}`,
          csv: match[2].trim()
        });
      }

      const validationMatch = text.match(/VALIDATION:([\s\S]*)$/);
      
      setTables(extractedTables);
      setValidation(validationMatch ? validationMatch[1].trim() : "Validation section missing in response.");
      
      if (extractedTables.length === 0) {
        setError("No tables were detected in the input provided.");
      }
    } catch (err) {
      console.error(err);
      setError(err instanceof Error ? err.message : "An unexpected error occurred during processing.");
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
        skipEmptyLines: false,
      });

      if (parseResult.data && parseResult.data.length > 0) {
        // Add spacing between tables (approx 1 inch gap = ~6 rows)
        if (index > 0) {
          combinedData.push([]);
          combinedData.push([]);
          combinedData.push([]);
          combinedData.push([]);
          combinedData.push([]);
          combinedData.push([]);
        }
        
        // Add a header row for the table name
        combinedData.push([`--- ${table.name} ---`]);
        
        // Add the table data, shifting "Subscription" cells to the right
        const processedRows = parseResult.data.map(row => {
          const newRow = [...row];
          // Iterate backwards to avoid shifting the same value multiple times
          for (let i = newRow.length - 1; i >= 0; i--) {
            if (newRow[i] && String(newRow[i]).toLowerCase().includes('subscription')) {
              const val = newRow[i];
              newRow[i] = ''; // Clear original
              newRow[i + 1] = val; // Move to right
            }
          }
          return newRow;
        });
        combinedData.push(...processedRows);
      }
    });

    if (combinedData.length > 0) {
      const ws = XLSX.utils.aoa_to_sheet(combinedData);
      
      // Set column widths to 200 pixels
      const maxCols = combinedData.reduce((max, row) => Math.max(max, row.length), 0);
      const maxRows = combinedData.length;
      ws['!cols'] = Array(maxCols).fill({ wpx: 200 });

      // Set row 2 height to 36 points (index 1 is row 2)
      if (!ws['!rows']) ws['!rows'] = [];
      ws['!rows'][1] = { hpt: 36 };

      // Apply styling to ALL cells
      const lightCreamFill = { fgColor: { rgb: "FFFDD0" } };
      const blueFill = { fgColor: { rgb: "0000FF" } };
      const boldFont = { sz: 11, bold: true, color: { rgb: "FFFFFF" } };
      const defaultFont = { sz: 10 };
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

      for (let r = 0; r < maxRows; r++) {
        const row = combinedData[r];
        
        // Check if this is a table name row to update active special columns
        if (row && row[0] && String(row[0]).startsWith('--- ') && String(row[0]).endsWith(' ---')) {
          activeCompanyNameCols.clear();
          activeTotalPriceCols.clear();
          activeCustomerNameCols.clear();
          
          // Peek at the next row for headers
          const headerRow = combinedData[r + 1];
          if (headerRow) {
            headerRow.forEach((cell, c) => {
              const cellText = String(cell).toLowerCase();
              if (cellText.includes('company name')) activeCompanyNameCols.add(c);
              if (cellText.includes('total price')) activeTotalPriceCols.add(c);
              if (cellText.includes('customer name')) activeCustomerNameCols.add(c);
            });
          }
        }

        for (let c = 0; c < maxCols; c++) {
          const cellAddress = XLSX.utils.encode_cell({ r, c });
          if (!ws[cellAddress]) {
            ws[cellAddress] = { t: 's', v: '' };
          }
          
          const isRow2 = r === 1;
          const isSpecialCol = activeCompanyNameCols.has(c) || activeTotalPriceCols.has(c) || activeCustomerNameCols.has(c);
          
          // Determine font
          let font: any = isRow2 ? { ...boldFont } : { ...defaultFont };
          if (isSpecialCol) {
            font.sz = 11;
            font.bold = true;
          }

          ws[cellAddress].s = {
            alignment: defaultAlignment,
            fill: isRow2 ? blueFill : lightCreamFill,
            font: font,
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
                  <p className="text-gray-500 text-sm max-w-xs mt-2">{error}</p>
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
