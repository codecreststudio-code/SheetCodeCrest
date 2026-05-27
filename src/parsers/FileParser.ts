export interface FileParser {
  /** Check if this parser can handle the given file */
  canParse(file: File): boolean | Promise<boolean>;

  /** Parse the file and return standardized data format */
  parse(file: File): Promise<{
    data: any[]; // DataRow[] equivalent
    originalRows: number;
    headerRow: number;
    sheetName: string;
    headers: string[];
  }>;
}