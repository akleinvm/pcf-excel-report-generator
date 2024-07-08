import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class ExcelReportGenerator implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _button: HTMLButtonElement;
    private _notifyOutputChanged: () => void;

    constructor(){}

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        this._button = document.createElement("button");
        this._button.addEventListener("click", this.generateReport.bind(this))
        this._notifyOutputChanged = notifyOutputChanged;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        
    }


    public getOutputs(): IOutputs
    {
        return {};
    }


    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }

    
    //Functions
    private generateReport(): void
    {

    }
}
