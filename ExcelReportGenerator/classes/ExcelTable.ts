import { CellIndex } from "../types";
import { ExcelColumnConverter } from "./ExcelColumnConverter";

export class ExcelTable {
    private _xmlDocument: Document;
    private _tableElement: Element;
    private _minCellIndex: CellIndex;
    private _maxCellIndex: CellIndex;
    private _autoFilterElement: Element;
  
    public fromXML(xmlString: string) {
      this._xmlDocument = new DOMParser().parseFromString(xmlString, "text/xml");
      this._tableElement = this._xmlDocument.getElementsByTagName('table')[0];
      const [minCellRef, maxCellRef] = this._tableElement.getAttribute('ref')?.split(':') ?? ['A1', 'A1'];
      this._minCellIndex = ExcelColumnConverter.cellRefToIndex(minCellRef);
      this._maxCellIndex = ExcelColumnConverter.cellRefToIndex(maxCellRef);
      this._autoFilterElement = this._tableElement.getElementsByTagName('autoFilter')[0];
    }
  
    public get minCellIndex(): CellIndex {
      return this._minCellIndex;
    }
  
    public get maxCellIndex(): CellIndex {
      return this._maxCellIndex;
    }
  
    public set maxCellIndex(cellIndex: CellIndex) {
      this._maxCellIndex = cellIndex;
    }
  
    public toString(): string {
      const minCellRef = ExcelColumnConverter.numberToColumn(this.minCellIndex.ColumnIndex) + this.minCellIndex.RowIndex;
      const maxCellRef = ExcelColumnConverter.numberToColumn(this.maxCellIndex.ColumnIndex) + this.maxCellIndex.RowIndex;
      const tableRef = `${minCellRef}:${maxCellRef}`;
  
      this._tableElement.setAttribute('ref', tableRef);
      this._autoFilterElement.setAttribute('ref', tableRef); 
  
      const xmlSerializer = new XMLSerializer();
      return xmlSerializer.serializeToString(this._xmlDocument);
    }
  }
  