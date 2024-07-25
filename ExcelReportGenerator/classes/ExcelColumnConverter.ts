import { CellIndex } from "../types";

export class ExcelColumnConverter {
    private static columnToNumberMap: Map<string, number> = new Map();
    private static numberToColumnMap: Map<number, string> = new Map();
  
    static cellRefToIndex(reference: string): CellIndex {
      const match = reference.match(/^([A-Z]+)(\d+)$/) ?? [];
      if(match.length < 3) throw new Error(`Invalid cell reference '${reference}'`);
      const [, columnLetter, rowNumber] = match ?? ['', ''];
  
      const rowIndex = Number(rowNumber);
      
      let columnIndex = 0;
      if (this.columnToNumberMap.has(columnLetter)) {
         columnIndex = this.columnToNumberMap.get(columnLetter) ?? -1;
      } else {
        columnIndex = 0;
        for (let i = 0; i < columnLetter.length; i++) {
          columnIndex = columnIndex * 26 + (columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
        }
  
        this.columnToNumberMap.set(columnLetter, columnIndex);
        
        if (!this.numberToColumnMap.has(columnIndex)) {
          this.numberToColumnMap.set(columnIndex, columnLetter);
        }
      }
  
      return {RowIndex: rowIndex, ColumnIndex: columnIndex}
    }
  
    static numberToColumn(number: number): string {
        if (this.numberToColumnMap.has(number)) {
            return this.numberToColumnMap.get(number)!;
        }
  
        let column = '';
        while (number > 0) {
            number--;  // Adjust for 0-indexing
            const remainder = number % 26;
            column = String.fromCharCode(remainder + 'A'.charCodeAt(0)) + column;
            number = Math.floor(number / 26);
        }
  
        this.numberToColumnMap.set(number, column);
  
        if (!this.columnToNumberMap.has(column)) {
          this.columnToNumberMap.set(column, number);
        }
  
        return column;
    }
}