import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as ExcelJS from 'exceljs';

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
        const contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

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
        
        const blob = this.convertBase64ToBlob(base64String, contentType);
        if (!blob || blob.size === 0) {
            console.error("Failed to create Blob from base64 string.");
            alert("Failed to create the file");
        }

        const workbook = this.convertBlobToWorkbook(blob);
        const updatedWorkbook = this.addRecordsToTable(await workbook, "", "");
        const updatedBlob = this.convertWorkbookToBlob(await updatedWorkbook);

        this.downloadBlob(await updatedBlob, fileName);

/*
        try {
            await this.generateExcelReport(fileName, payloadJsonString);
        } catch (error) {
            console.error("Error generating Excel report:", error);
            alert("Failed to generate the Excel report:" + (error instanceof Error ? error.message: String(error)));
        }*/
        
    }

    private downloadBlob(blob: Blob, fileName: string) {
        const blobUrl = URL.createObjectURL(blob);

        const downloadLink = document.createElement('a');
        downloadLink.href = blobUrl;
        downloadLink.download = fileName;

        this._container.appendChild(downloadLink);

        downloadLink.click();

        this._container.removeChild(downloadLink);
        URL.revokeObjectURL(blobUrl)
    }

    private async convertWorkbookToBlob(workbook: ExcelJS.Workbook): Promise<Blob> {
        const contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        const blob = await workbook.xlsx.writeBuffer();
        return new Blob([blob], {type: contentType})
    }

    private async convertBlobToWorkbook(blob: Blob): Promise<ExcelJS.Workbook> {
        const workbook = new ExcelJS.Workbook();

        const arrayBuffer = await blob.arrayBuffer();
        await workbook.xlsx.load(arrayBuffer);

        return workbook
    }

    private async addRecordsToTable(workbook: ExcelJS.Workbook, tableName: string, newRecords: string): Promise<ExcelJS.Workbook> {
        return workbook
    }

    private async generateExcelReport(workbook: ExcelJS.Workbook, payloadJsonString: string): Promise<void>
    {
        try {
            const updatedBlob = await this.addRecordsToExcelTable(workbook, tableName, payloadJsonString);

            const blobUrl = URL.createObjectURL(updatedBlob);

            const downloadLink = document.createElement('a');
            downloadLink.href = blobUrl;
            downloadLink.download = fileName;

            this._container.appendChild(downloadLink);

            downloadLink.click();

            this._container.removeChild(downloadLink);
            //URL.revokeObjectURL(blobUrl)
        } catch(error) {
            console.error("Error in generateExcelReport:", error);
            throw error;
        }
    }
    
    private convertBase64ToBlob(base64String: string, contentType: string) {
        const binaryString = atob(base64String);
        const uint8Array = Uint8Array.from(binaryString, char => char.charCodeAt(0));

        return new Blob([uint8Array], {type: contentType})
    }

    private async addRecordsToExcelTable(workbook: ExcelJS.Workbook, tableName: string, newRowsJson: string): Promise<Blob>
    {
        const worksheet = workbook.addWorksheet("My Sheet");

        worksheet.columns = [
            {header: 'Id', key: 'id', width: 10},
            {header: 'Name', key: 'name', width: 32}, 
            {header: 'D.O.B.', key: 'dob', width: 15,}
          ];
          
        worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970, 1, 1)});
        worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965, 1, 7)});

/*
        const worksheet = workbook.getWorksheet(1);
        if(!worksheet) {
            throw new Error("Worksheet not found");
        }

        const table = worksheet?.getTable(tableName);
        if(!table) {
            throw new Error(`Table with name ${tableName} not found`);
        }

        let newRows: any[];
        try {
            newRows = JSON.parse(newRowsJson);
            console.log("Parsed newRows:", newRows);

            if(!Array.isArray(newRows)) {
                throw new Error("newRows is not an array");
            }
        } catch(error) {
            console.error("Error parsing newRowsJson:", error);
            throw new Error(`Invalid JSON for newRows: ${error}`);
        }

        if(newRows.length === 0) {
            console.warn("newRows is empty");
        }

        newRows.forEach((newRow, index) => {
            if(newRow === null || typeof newRow !== 'object') {
                console.warn(`Invalid row at index ${index}:`, newRow);
                return;
            }
            try {
                table.addRow(Object.values(newRow));
            } catch(error) {
                console.error(`Error adding row at index ${index}:`, error);
            }
        });
*/

        console.log("Finished adding rows");
        const updatedWorkbookBlob = await workbook.xlsx.writeBuffer();
        console.log("Workbook updated and converted to blob");
        return new Blob([updatedWorkbookBlob], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
    }
}
