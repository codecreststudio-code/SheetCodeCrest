import * as XLSX from "xlsx";
import { FileParser } from "./FileParser";
import { DataRow } from "../excelAnalytics"; // Note: we will later change this to a common location or keep as is for now

// We are going to reuse the existing helper functions from excelAnalytics.ts.
// However, to avoid circular dependencies and to keep the parser self-contained,
// we will copy the necessary helper functions here or import them from a shared location.
// For simplicity and to avoid breaking changes, we will import from excelAnalytics.ts
// and then later we can refactor to share common utilities.

// But note: we are planning to rename excelAnalytics.ts to dataAnalytics.ts.
// For now, we will import from the existing file and adjust later.

import {
  detectHeaderRowForSheet,
  normalizeHeaders,
  buildTable,
  chooseBestSheet,
  parseWorkbook
} from "../excelAnalytics";

export class ExcelParser implements FileParser {
  async canParse(file: File): Promise<boolean> {
    const extension = file.name.split('.').pop()?.toLowerCase();
    return extension === 'xls' || extension === 'xlsx';
  }

  async parse(file: File): Promise<{
    data: DataRow[];
    originalRows: number;
    headerRow: number;
    sheetName: string;
    headers: string[];
  }> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target!.result, { type: "array" });
          const result = parseWorkbook(wb);
          resolve(result);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }
}