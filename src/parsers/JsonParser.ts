import { FileParser } from "./FileParser";
import { detectHeaderRowForSheet, buildTable } from "../excelAnalytics";

export class JsonParser implements FileParser {
  async canParse(file: File): Promise<boolean> {
    const extension = file.name.split(".").pop()?.toLowerCase();
    return extension === "json";
  }

  async parse(file: File): Promise<{
    data: any[];
    originalRows: number;
    headerRow: number;
    sheetName: string;
    headers: string[];
  }> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const text = e.target!.result as string;
          const json = JSON.parse(text);

          let target = json;

          // If the top level is an object, look for any property that contains an array
          if (typeof json === "object" && json !== null && !Array.isArray(json)) {
            // Check if there is an array property like "data", "records", "rows", etc.
            // Or just any array property
            let foundArray: any[] | undefined;
            for (const key in json) {
              if (Array.isArray(json[key])) {
                foundArray = json[key];
                break;
              }
            }
            if (foundArray) {
              target = foundArray;
            }
          }

          const sheetName = file.name.replace(/\.[^/.]+$/, ""); // remove extension

          if (Array.isArray(target)) {
            if (target.length > 0) {
              if (Array.isArray(target[0])) {
                // Case 1: Array of arrays (tabular)
                const raw = target as any[][];
                const headerRow = detectHeaderRowForSheet(raw, 12);
                const { headers, data } = buildTable(raw, headerRow);
                resolve({
                  data,
                  originalRows: raw.length,
                  headerRow,
                  sheetName,
                  headers,
                });
                return;
              } else if (typeof target[0] === "object" && target[0] !== null) {
                // Case 2: Array of objects
                // Standardize by extracting headers from keys of the objects
                const headersSet = new Set<string>();
                target.forEach((obj) => {
                  if (typeof obj === "object" && obj !== null) {
                    Object.keys(obj).forEach((k) => headersSet.add(k));
                  }
                });
                const headers = Array.from(headersSet);
                resolve({
                  data: target,
                  originalRows: target.length,
                  headerRow: 0,
                  sheetName,
                  headers,
                });
                return;
              }
            }

            // Empty array or other array type
            resolve({
              data: [],
              originalRows: 0,
              headerRow: 0,
              sheetName,
              headers: [],
            });
            return;
          }

          // Case 3: Single object (flatten it to a single row)
          const flattenObject = (obj: any, prefix = ""): any => {
            return Object.keys(obj).reduce((acc: any, key: string) => {
              const pre = prefix.length ? `${prefix}.` : "";
              if (typeof obj[key] === "object" && obj[key] !== null && !Array.isArray(obj[key])) {
                Object.assign(acc, flattenObject(obj[key], pre + key));
              } else {
                acc[pre + key] = obj[key];
              }
              return acc;
            }, {} as any);
          };

          if (typeof json === "object" && json !== null) {
            const flattened = flattenObject(json);
            const headers = Object.keys(flattened);
            resolve({
              data: [flattened],
              originalRows: 1,
              headerRow: 0,
              sheetName,
              headers,
            });
            return;
          }

          throw new Error("Invalid JSON structure");
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = () => reject(new Error("Failed to read file"));
      reader.readAsText(file);
    });
  }
}