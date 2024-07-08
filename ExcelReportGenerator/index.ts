import { IInputs, IOutputs } from "./generated/ManifestTypes";

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
    private onClick(): void 
    {
        const fileName = "export.xlsx";
        const base64String = this._context.parameters.Template.raw ?? "";

        if (!base64String) {
            console.error("Base64 string is empty");
            alert("No data available for export");
            return;
        }
        
        try {
            this.generateExcelReport(fileName, base64String);
        } catch (error) {
            console.error("Error generating Excel report:", error);
            alert("Failed to generate the Excel report");
        }
        
    }

    private generateExcelReport(fileName: string, base64String: string): void
    {
        const contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        const blob = this.base64ToBlob(base64String, contentType);

        if (!blob || blob.size === 0) {
            console.error("Failed to create Blob from base64 string.");
            alert("Failed to create the file");
        }

        const blobUrl = URL.createObjectURL(blob);

        const downloadLink = document.createElement('a');
        downloadLink.href = blobUrl;
        downloadLink.download = fileName;

        this._container.appendChild(downloadLink);

        downloadLink.click();

        this._container.removeChild(downloadLink);
        URL.revokeObjectURL(blobUrl);

    }
    
    private base64ToBlob(base64String: string, contentType: string) {
        const binaryString = atob(base64String);
        const uint8Array = Uint8Array.from(binaryString, char => char.charCodeAt(0));

        return new Blob([uint8Array], {type: contentType})
    }
}
