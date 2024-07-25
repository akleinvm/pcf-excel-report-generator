import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { convertBase64ToArrayBuffer, downloadBlob } from "./utils";
import { ExcelSharedStrings } from "./classes/ExcelSharedStrings";
import { ExcelStyles } from "./classes/ExcelStyles";
import { ExcelTable } from "./classes/ExcelTable";
import { ExcelWorksheet } from "./classes/ExcelWorksheet";
import { PowerAppsObject, PowerAppsArray, CellObject, CellFormat } from "./types";
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

  private async onClick(): Promise<void> 
  {
      console.log('Initializing constants...');
      const fileName = "export.xlsx";
      const sheetNo = this._context.parameters.SheetNo.raw ?? 1;

      console.log('Parsing template base64 string...');
      const base64String: string = this._context.parameters.Template.raw ?? '';
      if(base64String === '') throw new Error("template is null or undefined");

      const zipBuffer: ArrayBuffer = convertBase64ToArrayBuffer(base64String);
      
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
      const sharedStrings = new ExcelSharedStrings()
      sharedStrings.fromXML(xmlSharedStrings);

      console.log('Parsing styles.xml...');
      const xmlStyles = files.get("xl/styles.xml") ?? '';
      if(xmlStyles === '') throw new Error('styles is null or undefined');
      const styles = new ExcelStyles();
      styles.fromXML(xmlStyles);
      

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
      const payloadString: string = this._context.parameters.Payload.raw ?? '';
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

      const tableNo = this._context.parameters.TableNo.raw ?? 0;
      let startRowIndex = this._context.parameters.StartRowIndex.raw ?? 1;
      let startColumnIndex = this._context.parameters.StartColumnIndex.raw ?? 1;
      const table = new ExcelTable();
      if(tableNo > 0) {
        console.log(`Parsing table${tableNo}.xml...`);
        const xmlTable = files.get(`xl/tables/table${tableNo}.xml`) ?? '';
        table.fromXML(xmlTable); console.log(table);

        startRowIndex = table.maxCellIndex.RowIndex;
        startColumnIndex = table.minCellIndex.RowIndex;
        
        console.log('Updating table boundary...');
        table.maxCellIndex = {RowIndex: startRowIndex + payload.length, ColumnIndex: table.maxCellIndex.ColumnIndex};

        files.set(`xl/tables/table${tableNo}.xml`, table.toString());
      }

      console.log(`Parsing sheet${sheetNo}.xml...`);
      const xmlWorksheet = files.get(`xl/worksheets/sheet${sheetNo}.xml`) ?? "";
      if(xmlWorksheet === '') throw new Error('worksheet is null or undefined');
      const worksheet = new ExcelWorksheet();
      worksheet.fromXML(xmlWorksheet);

      console.log(`Adding rows to sheet${sheetNo}...`);
      worksheet.addRows(payload, startRowIndex, startColumnIndex);

      
      console.log('Updating file contents...');
      files.set("xl/sharedStrings.xml", sharedStrings.toString());
      files.set("xl/styles.xml", styles.toString());
      files.set(`xl/worksheets/sheet${sheetNo}.xml`, worksheet.toString());

      console.log('Generating blob...');
      const updatedZip = new JSZip();
      for(const [relativePath, file] of files) {
        updatedZip.file(relativePath, file);
      }

      const blob = updatedZip.generateAsync({type:"blob", compression: "DEFLATE", compressionOptions: {level: 9}})

      console.log('Downloading blob...');
      downloadBlob(fileName, await blob);

  }
}
