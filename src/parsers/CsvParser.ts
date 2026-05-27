import Papa from "papaparse";
import { FileParser } from "./FileParser";
import { detectHeaderRowForSheet, buildTable } from "../excelAnalytics";

export class CsvParser implements FileParser {
  async canParse(file: File): Promise<boolean> {
    const extension = file.name.split(".").pop()?.toLowerCase();
    return extension === "csv" || extension === "tsv";
  }

  async parse(file: File): Promise<{
    data: any[];
    originalRows: number;
    headerRow: number;
    sheetName: string;
    headers: string[];
  }> {
    return new Promise((resolve, reject) => {
      file.text()
        .then((text) => {
          const raw = Papa.parse(text, { header: false, skipEmptyLines: true }).data as string[][];
          if (!raw.length) {
            reject(new Error("CSV file is empty"));
            return;
          }

          // Now we have the raw 2D array (array of arrays of strings).
          const headerRow = detectHeaderRowForSheet(raw, 12);
          const { headers, data } = buildTable(raw, headerRow);

          // We need to return the same format as the ExcelParser.
          const sheetName = file.name.replace(/\.[^/.]+$/, ""); // remove extension
          const originalRows = raw.length;

          resolve({
            data,
            originalRows,
            headerRow,
            sheetName,
            headers,
          });
        })
        .catch(reject);
    });
  }
}