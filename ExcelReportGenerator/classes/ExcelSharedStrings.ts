export class ExcelSharedStrings {
    private _xmlDocument: XMLDocument;
    private _sstElement: Element;
    private _namespace: string;
    private _stringsMap: Map<string, number>;

    public fromXML(xmlString: string) {
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

    public toString(): string {
        this._sstElement.setAttribute('uniqueCount', this._stringsMap.size.toString());

        const xmlSerializer = new XMLSerializer();
        return xmlSerializer.serializeToString(this._xmlDocument);
    }

    public getStringIndex(string: string): number {
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