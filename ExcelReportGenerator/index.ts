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
        
        const zipBuffer = this.convertBase64ToArrayBuffer(base64String);
        
        const test = new ExcelSharedStrings();
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
        sharedStrings.addString("this is added just now");

        files.set("xl/sharedStrings.xml", sharedStrings.toXML());

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
  private _count: number;
  private _uniqueCount: number;
  private _stringsMap: Map<string, number>;

  constructor(count: number = 0, uniqueCount: number = 0, strings: string[] = []) {
    this._count = count;
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
    const xmlDoc = parser.parseFromString(xmlString, 'text/xml');

    const sstElement = xmlDoc.documentElement;
    const count = parseInt(sstElement.getAttribute('count') || '0', 10);
    const uniqueCount = parseInt(sstElement.getAttribute('uniqueCount') || '0', 10);

    this._count = count;
    this._uniqueCount = uniqueCount;
    
    const siElements = xmlDoc.getElementsByTagName('si');
    this._stringsMap.clear();
    for(let i=0; i<siElements.length; i++) {
      const tElement = siElements[i].getElementsByTagName('t')[0];
      if(tElement && tElement.textContent) {
        this._stringsMap.set(tElement.textContent, i)
      }
    }
  }

  toXML(): string {
    const xmlHeaderString = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const doc = new DOMParser().parseFromString('<sst/>', 'text/xml');

    const sstElement = doc.documentElement;

    this._stringsMap.forEach((index, text) => {
      const siElement = doc.createElementNS('', 'si',);
      const tElement = doc.createElement('t');
      tElement.textContent = text;
      siElement.appendChild(tElement);
      sstElement.appendChild(siElement);
    });

    sstElement.setAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    sstElement.setAttribute('count', this._count.toString());
    sstElement.setAttribute('uniqueCount', this._uniqueCount.toString());

    const xmlSerializer = new XMLSerializer();
    return xmlHeaderString + xmlSerializer.serializeToString(doc)
  }

  addString(string: string): void {
    if(!this._stringsMap.has(string)) {
      this._stringsMap.set(string, this._stringsMap.size);
      this._uniqueCount++
    }
  }

  getStringIndex(string: string): number {
    return this._stringsMap.get(string) ?? -1
  }

  setCount(count: number): void {
    this._count = count
  }
}