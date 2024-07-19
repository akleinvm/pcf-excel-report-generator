import { profileEnd } from "console";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as ExcelJS from 'exceljs';
import * as JSZip from 'jszip';

interface PowerAppsObject {Value: string}
interface PowerAppsArray {Value: Array<PowerAppsObject>}

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

      console.profile();
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
      
      const columnTypes: Array<string> = JSON.parse(this._context.parameters.ColumnTypes.raw ?? "");

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
      const sharedStrings = new ExcelSharedStrings(xmlSharedStrings);
      
      const payload: Array<PowerAppsArray> = JSON.parse(payloadJsonString);
      const array: Array<Array<number>> = [];
      const excelEpochTime = new Date(Date.UTC(1899, 11, 30)).getTime();
      const daysToMilliseconds = 86400000;
      for(let i = 0; i < payload.length; i++) {
        const row = payload[i];
        const cells: Array<number> = [];

        for(let j = 0; j < row.Value.length; j++) {
          const cellContent = row.Value[j].Value;
          let value: number;
          switch(columnTypes[j]) {
            case "String": value = sharedStrings.getStringIndex(cellContent); break;
            case "Number": value = Number(cellContent); break;
            case "Date": value = (new Date(cellContent).getTime() - excelEpochTime) / daysToMilliseconds + 1; break;
            case "Boolean": value = Number(cellContent === 'true'); break;
            default: value = sharedStrings.getStringIndex(cellContent); break;
          }
          cells.push(value);
        }
        array.push(cells);
      }

      const xmlWorksheet = files.get(`xl/worksheets/sheet${sheetNo}.xml`) ?? "";
      const worksheet = new ExcelWorksheet(xmlWorksheet);

      worksheet.addRows(array, columnTypes, 15, 15);
        
      files.set("xl/sharedStrings.xml", sharedStrings.toString());
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
}


class ExcelSharedStrings {
  private _xmlDocument: XMLDocument;
  private _sstElement: Element;
  private _namespace: string;
  private _count: number;
  private _stringsMap: Map<string, number>;

  constructor(xmlString: string) {
    this._xmlDocument = new DOMParser().parseFromString(xmlString, 'text/xml');
    this._sstElement = this._xmlDocument.getElementsByTagName('sst')[0];
    this._namespace = this._sstElement.getAttribute('xmlns') ?? '';
    this._count = Number(this._sstElement.getAttributeNS(this._namespace, 'count'));
    this._stringsMap = new Map();
    
    const siElements = this._sstElement.getElementsByTagName('si');
    for(let i = 0; i < siElements.length; i++) {
      const tElement = siElements[i].getElementsByTagName('t')[0];
      if(tElement && tElement.textContent) {
        this._stringsMap.set(tElement.textContent, i)
      }
    }
  }

  toString(): string {
    this._sstElement.setAttribute('count', this._count.toString());
    this._sstElement.setAttribute('uniqueCount', this._stringsMap.size.toString());

    const xmlSerializer = new XMLSerializer();
    return xmlSerializer.serializeToString(this._xmlDocument);
  }

  getStringIndex(string: string): number {
    let stringIndex = this._stringsMap.get(string);

    if(!stringIndex) {
      stringIndex = this._stringsMap.size;
      this._stringsMap.set(string, this._stringsMap.size);
      const siElement = this._xmlDocument.createElementNS(this._namespace, 'si',);
      const tElement = this._xmlDocument.createElementNS(this._namespace, 't');
      tElement.textContent = string;
      siElement.appendChild(tElement);
      this._sstElement.appendChild(siElement);
    }

    return stringIndex
  }

  setCount(count: number): void {
    this._count = count
  }

  incrementCount(): number {
    return this._count++
  }
}

interface ColumnFormat {Type: string, Attribute: string | null}

class ExcelWorksheet {
  private _xmlDoc: Document;
  private _worksheet: Element;
  private _namespace: string;
  private _sheetData: Element;

  private _dimension: Element;
  private _minRowIndex: number;
  private _minColumnIndex: number;
  private _maxRowIndex: number;
  private _maxColumnIndex: number;

  private _rowsMap: Map<number, Element>;
  private _cellsMap: Map<string, Element>;
  private _columnTypesMap: Map<string, ColumnFormat>; 

  constructor(xmlString: string) {
    this._xmlDoc = new DOMParser().parseFromString(xmlString, "text/xml");

    this._worksheet = this._xmlDoc.getElementsByTagName('worksheet')[0];
    this._namespace = this._worksheet.getAttribute('xmlns') ?? "";
    
    this._dimension = this._xmlDoc.getElementsByTagName("dimension")[0];
    const dimensionRef = this._dimension.getAttribute('ref') ?? '';
    const dimensionIndexes = dimensionRef.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/) ?? ['', '1', '1', '1', '1'];
    this._minRowIndex = Number(dimensionIndexes[2]);
    this._minColumnIndex = ExcelColumnConverter.columnToNumber(dimensionIndexes[1]);
    this._maxRowIndex = Number(dimensionIndexes[4]);
    this._maxColumnIndex = ExcelColumnConverter.columnToNumber(dimensionIndexes[3]);

    this._sheetData = this._xmlDoc.getElementsByTagName("sheetData")[0];
    this._rowsMap = new Map();
    const rows = this._sheetData.getElementsByTagName('row');
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      this._rowsMap.set(Number(row.getAttribute('r')), row);
    }

    this._cellsMap = new Map();
    const cells = this._sheetData.getElementsByTagName('c');
    for (let i = 0; i < cells.length; i++) {
      const cell = cells[i];
      this._cellsMap.set(cell.getAttribute('r') ?? '', cell);
    }

    this._columnTypesMap = new Map<string, ColumnFormat>([
      ['String', {Type: "s", Attribute: "t"}],
      ['Boolean', {Type: "b", Attribute: "t"}],
      ['Date', {Type: "5", Attribute: "s"}],
      ['Number', {Type: "n", Attribute: null}],
    ])

  }

  public addRows(rows: Array<Array<number>>, columnTypes: Array<string>, rowStartIndex: number, columnStartIndex: number): void {
    for(let i = 0; i < rows.length; i++) {
      const rowIndex = rowStartIndex + i;
      const row = rows[i];

      let rowElement = this._rowsMap.get(rowIndex);
      if(!rowElement) {
        rowElement = this._xmlDoc.createElementNS(this._namespace, 'row');
        rowElement.setAttribute('r', rowIndex.toString());
        rowElement.setAttribute('spans', `${columnStartIndex}:${columnStartIndex + row.length}`);
        rowElement.setAttribute('x14ac:dyDescent', "0.3");
        this._sheetData.appendChild(rowElement);
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
        
        const columnFormat = this._columnTypesMap.get(columnTypes[j]) ?? {Type: "s", Attribute: "t"};
        if(columnFormat.Attribute != null) cellElement.setAttribute(columnFormat.Attribute, columnFormat.Type);
        
        const valueElement = this._xmlDoc.createElementNS(this._namespace, 'v');
        valueElement.textContent = row[j].toString();
        cellElement.replaceChildren(valueElement);
        rowElement.appendChild(cellElement);
      }
    }
    this._maxRowIndex = Math.max(this._maxRowIndex, rowStartIndex + rows.length);
    this._maxColumnIndex = Math.max(this._maxColumnIndex, columnStartIndex + columnTypes.length - 1);
  }

  public toString(): string {
    const minDimensionRef = ExcelColumnConverter.numberToColumn(this._minColumnIndex) + this._minRowIndex;
    const maxDimensionRef = ExcelColumnConverter.numberToColumn(this._maxColumnIndex) + this._maxRowIndex;
    this._dimension.setAttribute('ref', `${minDimensionRef}:${maxDimensionRef}`);
    return new XMLSerializer().serializeToString(this._xmlDoc);
  }
}

class ExcelColumnConverter {
  private static columnToNumberMap: Map<string, number> = new Map();
  private static numberToColumnMap: Map<number, string> = new Map();

  static columnToNumber(column: string): number {
      // Check if the column number is already calculated
      if (this.columnToNumberMap.has(column)) {
          return this.columnToNumberMap.get(column)!;
      }

      let number = 0;
      for (let i = 0; i < column.length; i++) {
        number = number * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
      }

      // Store the calculated column number in the map
      this.columnToNumberMap.set(column, number);
      
      if (!this.numberToColumnMap.has(number)) {
        this.numberToColumnMap.set(number, column);
      }

      return number;
  }

  static numberToColumn(number: number): string {
      // Check if the column letter is already calculated
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

      // Store the calculated column letter in the map
      this.numberToColumnMap.set(number, column);

      if (!this.columnToNumberMap.has(column)) {
        this.columnToNumberMap.set(column, number);
      }

      return column;
  }
}