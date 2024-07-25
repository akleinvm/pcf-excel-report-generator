import { CellObject } from "../types";
import { ExcelColumnConverter } from "./ExcelColumnConverter";

export class ExcelWorksheet {
  private _xmlDoc: Document;
  private _worksheetElement: Element;
  private _namespace: string;
  private _sheetDataElement: Element;

  private _rowsMap: Map<number, Element>;
  private _cellsMap: Map<string, Element>;

  public fromXML(xmlString: string) {
    this._xmlDoc = new DOMParser().parseFromString(xmlString, "text/xml");

    this._worksheetElement = this._xmlDoc.getElementsByTagName('worksheet')[0];
    this._namespace = this._worksheetElement.getAttribute('xmlns') ?? "";

    this._sheetDataElement = this._xmlDoc.getElementsByTagName("sheetData")[0];
    this._rowsMap = new Map();
    const rows = this._sheetDataElement.getElementsByTagName('row');
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      this._rowsMap.set(Number(row.getAttribute('r')), row);
    }

    this._cellsMap = new Map();
    const cells = this._sheetDataElement.getElementsByTagName('c');
    for (let i = 0; i < cells.length; i++) {
      const cell = cells[i];
      this._cellsMap.set(cell.getAttribute('r') ?? '', cell);
    }
  }

  public addRows(rows: Array<Array<CellObject>>, rowStartIndex: number, columnStartIndex: number): void {
    for(let i = 0; i < rows.length; i++) {
      const rowIndex = rowStartIndex + i;
      const row = rows[i];

      let rowElement = this._rowsMap.get(rowIndex);
      if(!rowElement) {
        rowElement = this._xmlDoc.createElementNS(this._namespace, 'row');
        rowElement.setAttribute('r', rowIndex.toString());
        rowElement.setAttribute('spans', `${columnStartIndex}:${columnStartIndex + row.length}`);
        this._sheetDataElement.appendChild(rowElement);
        this._rowsMap.set(rowIndex, rowElement);
      } else {
        const [minSpan, maxSpan] = rowElement.getAttribute('spans')?.split(':') ?? [];
        const spans = `${Math.min(columnStartIndex, Number(minSpan)).toString()}:${Math.max(columnStartIndex + row.length - 1, Number(maxSpan))}`;
        rowElement.setAttribute('spans', spans);
      }
      
      for(let j = 0; j < row.length; j++) {
        const columnIndex = columnStartIndex + j;
        
        const cellReference = ExcelColumnConverter.numberToColumn(columnIndex) + rowIndex;
        let cellElement = this._cellsMap.get(cellReference);
        if(!cellElement) {
          cellElement = this._xmlDoc.createElementNS(this._namespace, 'c');
          cellElement.setAttribute('r', cellReference);
        }
        
        const cellStyle = row[j].Format.Style;
        if(cellStyle) cellElement.setAttribute('s', cellStyle ?? "");

        const cellType = row[j].Format.Type;
        if(cellType) cellElement.setAttribute('t', cellType ?? "");
        
        const valueElement = this._xmlDoc.createElementNS(this._namespace, 'v');
        valueElement.textContent = row[j].Value.toString();
        cellElement.replaceChildren(valueElement);
        rowElement.appendChild(cellElement);
      }
    }
  }

  public toString(): string {
    return new XMLSerializer().serializeToString(this._xmlDoc);
  }
}