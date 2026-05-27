import { FileParser } from "./FileParser";
import { ExcelParser } from "./ExcelParser";
import { CsvParser } from "./CsvParser";
import { JsonParser } from "./JsonParser";

export class ParserRegistry {
  private parsers: FileParser[];

  constructor() {
    this.parsers = [
      new ExcelParser(),
      new CsvParser(),
      new JsonParser()
    ];
  }

  /** Register a new parser */
  registerParser(parser: FileParser) {
    this.parsers.push(parser);
  }

  /** Get the first parser that can parse the given file */
  async getParserForFile(file: File): Promise<FileParser | null> {
    for (const parser of this.parsers) {
      if (await parser.canParse(file)) {
        return parser;
      }
    }
    return null;
  }

  /** Parse the file using the appropriate parser */
  async parseFile(file: File): Promise<{
    data: any[];
    originalRows: number;
    headerRow: number;
    sheetName: string;
    headers: string[];
  }> {
    const parser = await this.getParserForFile(file);
    if (!parser) {
      throw new Error(`No parser found for file: ${file.name}`);
    }
    return parser.parse(file);
  }
}

// Export a singleton instance for convenience
export const parserRegistry = new ParserRegistry();