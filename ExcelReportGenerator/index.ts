import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as ExcelJS from 'exceljs';
import * as JSZip from 'jszip';

export class ExcelReportGenerator implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _button: HTMLButtonElement;
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;

    constructor(){}

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        this._notifyOutputChanged = notifyOutputChanged;
        this._context = context;

        this._button = document.createElement("button");
        this._button.className = "primary-button";
        this._button.innerText = "Export";
        this._button.addEventListener("click", this.onClick.bind(this));

        this._container = document.createElement("div");
        this._container.appendChild(this._button);

        container.appendChild(this._container);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        this._context = context;
    }


    public getOutputs(): IOutputs
    {
        return {};
    }


    public destroy(): void
    {
        this._button.removeEventListener("click", this.onClick);
    }

    
    //Functions
    private async onClick(): Promise<void> 
    {
        const fileName = "export.xlsx";
        const zipMimeType = "application/zip";

        const base64String = this._context.parameters.Template.raw;
        if (!base64String) {
            console.error("Base64 string is empty");
            alert("No data available for export");
            return;
        }
        
        const payloadJsonString = this._context.parameters.Payload.raw;
        if (!payloadJsonString) {
            console.error("Payload JSON string is empty");
            alert("No payload data available for export");
            return;
        }

        const sheetNo = this._context.parameters.SheetNo.raw ?? 1;
        const startRowIndex = this._context.parameters.StartRowIndex.raw ?? 1;
        const startColumnIndex = this._context.parameters.StartColumnIndex.raw ?? 1;
        
        const zipBuffer = this.convertBase64ToArrayBuffer(base64String);
        
        const zip = new JSZip();
        await zip.loadAsync(zipBuffer);
        
        const files = new Map<string, string>();

        const zipFolder = zip.folder("");
        if(zipFolder) {
            for(const [relativePath, file] of Object.entries(zipFolder.files)) {
                const content = await file.async("binarystring");
                files.set(relativePath, content);
            }
        }

        const xmlSharedStrings = files.get("xl/sharedStrings.xml") ?? "";

        const sharedStrings = new ExcelSharedStrings();
        sharedStrings.fromXML(xmlSharedStrings);

        const xmlWorksheet = files.get(`xl/worksheets/sheet${sheetNo}.xml`) ?? "";
        const worksheet = new ExcelWorksheet(xmlWorksheet);

        const payload: Array<PowerAppsArray> = JSON.parse(payloadJsonString);

        payload.forEach((row, rowNo) => {
          row.Value.forEach((cell, columnNo) => {
            let sharedStringIndex: number | undefined;
            let type: string | undefined;
            let value: number = 0;

            if(cell.Value.charAt(0) === "'") {
              type = 's'
            } else if(!isNaN(Number(cell))) {
              value = Number(cell);
            } else if(cell.Value.toLowerCase() === 'true' || cell.Value.toLowerCase() === 'false') {
              type = 'b'
            } else if(!isNaN(new Date(cell.Value).getTime()) && cell.Value.length > 4) {
              type = '5'
              const date = new Date(cell.Value);
              const excelEpoch = new Date(Date.UTC(1899, 11, 30));
              const daysSinceExcelEpoch = (date.getTime() - excelEpoch.getTime()) / (24 * 60 * 60 * 1000);
              value = daysSinceExcelEpoch + 1;
            } else {
              type = 's'
            }

            if(type === 's') {
              if(cell.Value.charAt(0) === "'") {
                cell.Value = cell.Value.slice(1);
              }
              sharedStringIndex = sharedStrings.getStringIndex(cell.Value);
              if(sharedStringIndex === -1) {
                sharedStrings.addString(cell.Value);
                sharedStringIndex = sharedStrings.getStringIndex(cell.Value)
              }
              value = sharedStringIndex;
            }
            
            worksheet.addCell(rowNo + startRowIndex, columnNo + startColumnIndex, value, type);
          })
        })
          
        files.set("xl/sharedStrings.xml", sharedStrings.toXML());
        files.set("xl/worksheets/sheet2.xml", worksheet.toString());
        

        const updatedZip = new JSZip();
        for(const [relativePath, file] of files) {
          updatedZip.file(relativePath, file)
        }

        const blob = updatedZip.generateAsync({type:"blob", compression: "DEFLATE", compressionOptions: {level: 9}})
        this.downloadBlob("export.zip", await blob);

    }

    private convertBase64ToArrayBuffer(base64Content: string): ArrayBuffer {
        const binaryString = atob(base64Content);
        const bytes = Uint8Array.from(binaryString, char => char.charCodeAt(0));

        return bytes.buffer
    }

    private downloadBlob(fileName: string, blob: Blob): void {
        const blobUrl = URL.createObjectURL(blob);

        const downloadLink = document.createElement('a');
        downloadLink.href = blobUrl;
        downloadLink.download = fileName;

        this._container.appendChild(downloadLink);

        downloadLink.click();

        this._container.removeChild(downloadLink);
        URL.revokeObjectURL(blobUrl)
    }

    private getTypeCode(input: string): string {
      if(input === '' || input.charAt(0) === "'") return 's';
      if(input.toLowerCase() === 'true' || input.toLowerCase() === 'false') return 'b';

      const dateObject = new Date(input);
      if(!isNaN(dateObject.getTime()) && input.length > 4) return '5';

      return 's';

    }
}

interface PowerAppsObject {
  Value: string;
}

interface PowerAppsArray {
  Value: Array<PowerAppsObject>
}

class ExcelSharedStrings {
  private _count: number;
  private _stringsMap: Map<string, number>;
  private _xmlDocument: XMLDocument;

  constructor() {
    this._count = 0;
    this._stringsMap = new Map();
  }

  get count(): number {
    return this._count;
  }

  get uniqueCount(): number {
    return this._stringsMap.size;
  }

  get strings(): string[] {
    return Array.from(this._stringsMap.keys());
  }

  fromXML(xmlString: string): void {
    const parser = new DOMParser();
    this._xmlDocument = parser.parseFromString(xmlString, 'text/xml');

    const sstElement = this._xmlDocument.documentElement;
    const count = parseInt(sstElement.getAttribute('count') || '0', 10);

    this._count = count;
    
    const siElements = this._xmlDocument.getElementsByTagName('si');
    this._stringsMap.clear();
    for(let i=0; i<siElements.length; i++) {
      const tElement = siElements[i].getElementsByTagName('t')[0];
      if(tElement && tElement.textContent) {
        this._stringsMap.set(tElement.textContent, i)
      }
    }
  }

  toXML(): string {
    const xmlDocument = this._xmlDocument;
    const sstElement = xmlDocument.querySelectorAll('sst')[0];
    const xmlnsValue = sstElement.getAttribute('xmlns') ?? "";
    sstElement.replaceChildren();

    this._stringsMap.forEach((index, text) => {
      const siElement = xmlDocument.createElementNS(xmlnsValue, 'si',);
      const tElement = xmlDocument.createElementNS(xmlnsValue, 't');
      tElement.textContent = text;
      siElement.appendChild(tElement);
      sstElement.appendChild(siElement);
    });

    sstElement.setAttribute('count', this._count.toString());
    sstElement.setAttribute('uniqueCount', this._stringsMap.size.toString());

    const xmlSerializer = new XMLSerializer();
    return xmlSerializer.serializeToString(xmlDocument)
  }

  addString(string: string): number {
    if(!this._stringsMap.has(string)) {
      this._stringsMap.set(string, this._stringsMap.size);
    }

    return this._stringsMap.size
  }

  getStringIndex(string: string): number {
    return this._stringsMap.get(string) ?? -1
  }

  setCount(count: number): void {
    this._count = count
  }

  incrementCount(): number {
    return this._count++
  }
}

type CellData = {
  value: number;
  type?: string;
};

type RowData = Record<string, CellData>;

class ExcelWorksheet {
  private parser: DOMParser;
  private xmlDoc: Document;
  private namespace: string;
  private sheetData: Element;
  private dimension: Element;
  private rows: Record<number, RowData>;
  private maxColumn: string = 'A';
  private maxRow: number = 1;

  constructor(xmlString: string) {
    this.parser = new DOMParser();
    this.xmlDoc = this.parser.parseFromString(xmlString, "text/xml");
    this.namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    this.sheetData = this.xmlDoc.getElementsByTagNameNS(this.namespace, "sheetData")[0];
    this.dimension = this.xmlDoc.getElementsByTagNameNS(this.namespace, "dimension")[0];
    this.rows = {};
    this.parseExistingData();
  }

  private parseExistingData(): void {
    const rowElements = this.sheetData.getElementsByTagNameNS(this.namespace, "row");
    for (let i = 0; i < rowElements.length; i++) {
      const rowElement = rowElements[i];
      const rowIndex = parseInt(rowElement.getAttribute("r") || "0", 10);
      this.rows[rowIndex] = {};
      const cellElements = rowElement.getElementsByTagNameNS(this.namespace, "c");
      for (let j = 0; j < cellElements.length; j++) {
        const cellElement = cellElements[j];
        const cellReference = cellElement.getAttribute("r") || "";
        const columnLetter = cellReference.replace(/[0-9]/g, '');
        const value = parseInt(cellElement.getElementsByTagNameNS(this.namespace, "v")[0]?.textContent || '0', 10);
        const type = cellElement.getAttribute("t") || undefined;
        this.rows[rowIndex][columnLetter] = {value, type};
        this.updateMaxColumnAndRow(columnLetter, rowIndex);
      }
    }
    this.updateDimensionRef();
  }

  private updateMaxColumnAndRow(column: string, row: number): void {
    if (column > this.maxColumn) this.maxColumn = column;
    if (row > this.maxRow) this.maxRow = row;
  }

  private updateDimensionRef(): void {
    this.dimension.setAttribute("ref", `A1:${this.maxColumn}${this.maxRow}`);
  }

  private updateRowSpans(rowElement: Element, newColumn: string): void {
    const currentSpans = rowElement.getAttribute("spans") || "1:1";
    const [minCol, maxCol] = currentSpans.split(':').map(Number);
    const newMaxCol = Math.max(maxCol, this.columnLetterToNumber(newColumn));
    rowElement.setAttribute("spans", `${minCol}:${newMaxCol}`);
  }

  private columnLetterToNumber(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result *= 26;
      result += column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
    }
    return result;
  }

  private numberToColumnLetter(num: number): string {
    let result = '';
    while (num > 0) {
      num--;
      result = String.fromCharCode('A'.charCodeAt(0) + (num % 26)) + result;
      num = Math.floor(num / 26);
    }
    return result || 'A';
  }

  public addCell(rowIndex: number, columnIndex: number, value: number, type?: string): void {
    if (!this.rows[rowIndex]) {
      this.addRow(rowIndex);
    }
    const columnLetter = this.numberToColumnLetter(columnIndex);
    this.rows[rowIndex][columnLetter] = { value, type };

    let rowElement = this.sheetData.querySelector(`row[r="${rowIndex}"]`) as Element;
    if (!rowElement) {
      rowElement = this.xmlDoc.createElementNS(this.namespace, "row");
      rowElement.setAttribute("r", rowIndex.toString());
      this.sheetData.appendChild(rowElement);
    }

    let cellElement = rowElement.querySelector(`c[r="${columnLetter}${rowIndex}"]`);
    if (!cellElement) {
      cellElement = this.xmlDoc.createElementNS(this.namespace, "c");
      cellElement.setAttribute("r", `${columnLetter}${rowIndex}`);
      rowElement.appendChild(cellElement);
    }

    if (type) {
      cellElement.setAttribute("t", type);
    } else {
      cellElement.removeAttribute("t");
    }

    let vElement = cellElement.getElementsByTagNameNS(this.namespace, "v")[0];
    if (!vElement) {
      vElement = this.xmlDoc.createElementNS(this.namespace, "v");
      cellElement.appendChild(vElement);
    }
    vElement.textContent = value.toString();

    this.updateRowSpans(rowElement, columnLetter);
    this.updateMaxColumnAndRow(columnLetter, rowIndex);
    this.updateDimensionRef();
  }

  public addRow(rowIndex: number): void {
    if (!this.rows[rowIndex]) {
      this.rows[rowIndex] = {};
      const rowElement = this.xmlDoc.createElementNS(this.namespace, "row");
      rowElement.setAttribute("r", rowIndex.toString());
      rowElement.setAttribute("spans", "1:1");
      this.sheetData.appendChild(rowElement);
      this.updateMaxColumnAndRow(this.maxColumn, rowIndex);
      this.updateDimensionRef();
    }
  }

  public getCell(rowIndex: number, columnLetter: string): CellData | null {
    return this.rows[rowIndex]?.[columnLetter] || null;
  }

  public toString(): string {
    return new XMLSerializer().serializeToString(this.xmlDoc);
  }
}

/*
class ExcelWorksheet {
  private _xmlDocument: XMLDocument;

  fromXML(xmlString: string): void {
    const parser = new DOMParser();
    this._xmlDocument = parser.parseFromString(xmlString, 'text/xml');
  }

  toXML(): string {
    const xmlSerializer = new XMLSerializer();
    return xmlSerializer.serializeToString(this._xmlDocument);
  }

  addRecords(startCell: string, records: string[][]): void {
    const sheetData = this._xmlDocument.querySelectorAll('sheetData')[0];

    for(let row=1; row<=records.length; row++) {
      const currentRow = sheetData.querySelectorAll('row [r=' + row + ']')[0]

    }
  }
}*/

/*
class ExcelWorksheet {
  private _xmlDocument: XMLDocument;
  private _rowsMap: Map<string, Element>;
  private _cellArray: {rowNo: number, columnNo: number, reference: string, element: Element}[];

  private convertReferenceToIndexes(reference: string): Array<number> {
    reference.toUpperCase();
    reference.match("/^([A-Z]+)(\d+)$/");
    const [, letters, numbers] = reference;

    let rowNo = parseInt(numbers);
    let columnNo = 0;
    for(let i=0; i<letters.length; i++) {
      columnNo *= 26;
      columnNo += letters.charCodeAt(i) - 64;
    }

    return [rowNo, columnNo]; 
  }

  private convertIndexestoReference(rowNo: number, columnNo: number) {
    let result = '';
    while(columnNo > 0) {
      columnNo--;
      result = String.fromCharCode((columnNo % 26) + 65) + result;
      columnNo = Math.floor(columnNo / 26);
    }
    return result + rowNo;
  }

  fromXML(xmlString: string): void {
    this._rowsMap.clear();

    const parser = new DOMParser();
    this._xmlDocument = parser.parseFromString(xmlString, 'text/xml');

    const worksheet = this._xmlDocument.querySelectorAll('worksheet')[0];
    const sheetData = worksheet.querySelectorAll('sheetData')[0];

    const rows = sheetData.querySelectorAll('row');
    for(let i=0; i<rows.length; i++) {
      const row = rows[i];
      row.replaceChildren();
      const rowNo = row.getAttribute('r');

      if(row && rowNo) {
        this._rowsMap.set(rowNo, row)
      }
    }
    
    const cells = sheetData.querySelectorAll('c');
    for(let i=0; i<cells.length; i++) {
      const cell = cells[i];
      const reference = cell.getAttribute('r');

      if(cell && reference) {
        const [rowNo, columnNo] = this.convertReferenceToIndexes(reference);
        this._cellArray.push({rowNo: rowNo, columnNo: columnNo, reference: reference, element: cell});
      }
    }
  }
  /*
  fromXML(xmlString: string): void {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlString, 'text/xml');
    new XMLDocument().createElement("")
    const worksheetElement = xmlDoc.documentElement;

    worksheetElement.getElementsByTagName('row')
    worksheetElement.attributes.
  }
}*/