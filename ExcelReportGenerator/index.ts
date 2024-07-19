import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as JSZip from 'jszip';

interface PowerAppsObject {Value: string}
interface PowerAppsArray {Value: Array<PowerAppsObject>}
interface CellFormat {Type: string | null, Style: string | null}
interface CellObject {Value: number, Format: CellFormat}

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
      console.log('Initializing constants...');
      const fileName = "export.xlsx";
      const sheetNo = this._context.parameters.SheetNo.raw ?? 1;
      const startRowIndex = this._context.parameters.StartRowIndex.raw ?? 1;
      const startColumnIndex = this._context.parameters.StartColumnIndex.raw ?? 1;

      console.log('Parsing template base64 string...');
      const base64String: string = this._context.parameters.Template.raw ?? '';
      if(base64String === '') throw new Error("template is null or undefined");

      const zipBuffer: ArrayBuffer = this.convertBase64ToArrayBuffer(base64String);
      
      console.log('Retrieving files from template...');
      const zip = new JSZip();
      await zip.loadAsync(zipBuffer);
      
      const files = new Map<string, string>();

      const zipFolder = zip.folder("");
      if(zipFolder) {
          for(const [relativePath, file] of Object.entries(zipFolder.files)) {
              const content = await file.async('binarystring');
              files.set(relativePath, content);
          }
      }

      console.log('Parsing sharedStrings.xml...');
      const xmlSharedStrings = files.get('xl/sharedStrings.xml') ?? '';
      if(xmlSharedStrings === '') throw new Error('sharedStrings is null or undefined');
      const sharedStrings = new ExcelSharedStrings(xmlSharedStrings);

      console.log('Parsing styles.xml...');
      const xmlStyles = files.get("xl/styles.xml") ?? '';
      if(xmlStyles === '') throw new Error('styles is null or undefined');
      const styles = new ExcelStyles(xmlStyles);

      console.log('Generating columnTypes...');
      const columnTypesString: string = this._context.parameters.ColumnTypes.raw ?? '';
      const columnTypesPAObject: Array<PowerAppsObject> = JSON.parse(columnTypesString);

      const cellFormatsMap = new Map<string, CellFormat> ([
        ['String', {Type: 's', Style: null}],
        ['Number', {Type: null, Style: null}],
        ['Date', {Type: null, Style: styles.getFormatIndex(14)}],
        ['Boolean', {Type: 'b', Style: null}]
      ]);
      const cellFormatDefault: CellFormat = {Type: 's', Style: null};

      const columnTypes = new Array<string>;
      columnTypesPAObject.forEach(paObject => {
        columnTypes.push(paObject.Value);
      });
      
      console.log('Parsing payload...');
      const payloadString: string = this._context.parameters.Payload.raw ?? "";
      if(payloadString === '') throw new Error('payload is null or undefined');
      const payloadPAObject: Array<PowerAppsArray> = JSON.parse(payloadString); 
      const payload: Array<Array<CellObject>> = [];

      const excelEpochTime = new Date(Date.UTC(1899, 11, 30)).getTime();
      const millisecondsToDays = 1/86400000;
      for(let i = 0; i < payloadPAObject.length; i++) {
        const row = payloadPAObject[i];
        const cells: Array<CellObject> = [];

        for(let j = 0; j < row.Value.length; j++) {
          const cellContent = row.Value[j].Value;
          const cellType = columnTypes[j];
          let value: number;
          switch(cellType) {
            case "String": value = sharedStrings.getStringIndex(cellContent); break;
            case "Number": value = Number(cellContent); break;
            case "Date": value = (new Date(cellContent).getTime() - excelEpochTime) * millisecondsToDays + 1; break;
            case "Boolean": value = Number(cellContent === 'true'); break;
            default: value = sharedStrings.getStringIndex(cellContent); break;
          }
          cells.push({Value: value, Format: cellFormatsMap.get(cellType) ?? cellFormatDefault});
        }
        payload.push(cells);
      }

      console.log(`Parsing sheet${sheetNo}.xml...`);
      const xmlWorksheet = files.get(`xl/worksheets/sheet${sheetNo}.xml`) ?? "";
      if(xmlWorksheet === '') throw new Error('worksheet is null or undefined');
      const worksheet = new ExcelWorksheet(xmlWorksheet);

      console.log(`Adding rows to sheet${sheetNo}...`);
      worksheet.addRows(payload, startRowIndex, startColumnIndex);
      
      console.log('Updated file contents...');
      files.set("xl/sharedStrings.xml", sharedStrings.toString());
      files.set("xl/styles.xml", styles.toString());
      files.set("xl/worksheets/sheet2.xml", worksheet.toString());

      console.log('Generating blob...');
      const updatedZip = new JSZip();
      for(const [relativePath, file] of files) {
        updatedZip.file(relativePath, file);
      }

      const blob = updatedZip.generateAsync({type:"blob", compression: "DEFLATE", compressionOptions: {level: 9}})

      console.log('Downloading blob...');
      this.downloadBlob(fileName, await blob);

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
  private _stringsMap: Map<string, number>;

  constructor(xmlString: string) {
    this._xmlDocument = new DOMParser().parseFromString(xmlString, 'text/xml');
    this._sstElement = this._xmlDocument.getElementsByTagName('sst')[0];
    this._namespace = this._sstElement.getAttribute('xmlns') ?? '';
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
}

class ExcelStyles {
  private _xmlDocument: Document;
  private _styleSheet: Element;
  private _namespace: string;
  private _cellFormats: Element;
  private _cellFormatArray: Array<number>;

  constructor(xmlString: string) {
    this._xmlDocument = new DOMParser().parseFromString(xmlString, "text/xml"); 
    this._styleSheet = this._xmlDocument.getElementsByTagName('styleSheet')[0]; 
    this._namespace = this._styleSheet.getAttribute('xmlns') ?? ""; 

    this._cellFormats = this._styleSheet.getElementsByTagName('cellXfs')[0]; 
    const formats = this._cellFormats.getElementsByTagName('xf'); 
    
    this._cellFormatArray = [];
    for (let i = 0; i < formats.length; i++) {
      const format = formats[i];
      const formatId = Number(format.getAttribute('numFmtId'));
      this._cellFormatArray.push(formatId);
    }
    console.log(this._cellFormatArray);
  }

  getFormatIndex(formatId: number): string {
    let formatIndex = this._cellFormatArray.indexOf(formatId).toString();

    if(formatIndex === '-1') {
      formatIndex = this._cellFormatArray.length.toString();
      this._cellFormatArray.push(formatId);
      const xfElement = this._xmlDocument.createElementNS(this._namespace, 'xf');
      xfElement.setAttribute('numFmtId', formatId.toString());
      this._cellFormats.appendChild(xfElement);
    }

    return formatIndex
  }

  toString() {
    const xmlSerializer = new XMLSerializer();
    return xmlSerializer.serializeToString(this._xmlDocument);
  }
}

class ExcelWorksheet {
  private _xmlDoc: Document;
  private _worksheet: Element;
  private _namespace: string;
  private _sheetData: Element;

  private _rowsMap: Map<number, Element>;
  private _cellsMap: Map<string, Element>;

  constructor(xmlString: string) {
    this._xmlDoc = new DOMParser().parseFromString(xmlString, "text/xml");

    this._worksheet = this._xmlDoc.getElementsByTagName('worksheet')[0];
    this._namespace = this._worksheet.getAttribute('xmlns') ?? "";

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



class ExcelColumnConverter {
  private static columnToNumberMap: Map<string, number> = new Map();
  private static numberToColumnMap: Map<number, string> = new Map();

  static columnToNumber(column: string): number {
      if (this.columnToNumberMap.has(column)) {
          return this.columnToNumberMap.get(column)!;
      }

      let number = 0;
      for (let i = 0; i < column.length; i++) {
        number = number * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
      }

      this.columnToNumberMap.set(column, number);
      
      if (!this.numberToColumnMap.has(number)) {
        this.numberToColumnMap.set(number, column);
      }

      return number;
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