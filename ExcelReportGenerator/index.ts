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
        
        const zip = new JSZip();
        zip.loadAsync(zipBuffer);
        zip.generateAsync({type:"blob"}).then((base64) => )
        
        let files!: {[key: string]: string};

        zip.folder("")?.forEach(async function (relativePath, file) {
            console.log("This is a name: " + relativePath);
            const content = await file.async('string');
            files[relativePath] = content;
        })

        /*
        for (const [filename, file] of Object.entries(zip.files)) {
            console.log("This is a name: " + filename);
            const content = await file.async('string');
            files[filename] = content;
        }*/

        console.log(files["sharedStrings.xml"])

    }

    private convertBase64ToArrayBuffer(base64Content: string): ArrayBuffer {
        const binaryString = atob(base64Content);
        const bytes = Uint8Array.from(binaryString, char => char.charCodeAt(0));

        return bytes.buffer
    }
}
