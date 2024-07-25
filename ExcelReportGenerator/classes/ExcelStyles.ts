
export class ExcelStyles {
    private _xmlDocument: Document;
    private _styleSheetElement: Element;
    private _namespace: string;
    private _cellFormatsElement: Element;
    private _cellFormatArray: Array<number>;
  
    public fromXML(xmlString: string) {
      this._xmlDocument = new DOMParser().parseFromString(xmlString, "text/xml"); 
      this._styleSheetElement = this._xmlDocument.getElementsByTagName('styleSheet')[0]; 
      this._namespace = this._styleSheetElement.getAttribute('xmlns') ?? ""; 
  
      this._cellFormatsElement = this._styleSheetElement.getElementsByTagName('cellXfs')[0]; 
      const formats = this._cellFormatsElement.getElementsByTagName('xf'); 
      
      this._cellFormatArray = [];
      for (let i = 0; i < formats.length; i++) {
        const format = formats[i];
        const formatId = Number(format.getAttribute('numFmtId'));
        this._cellFormatArray.push(formatId);
      }
      console.log(this._cellFormatArray);
    }
  
    public getFormatIndex(formatId: number): string {
      let formatIndex = this._cellFormatArray.indexOf(formatId).toString();
  
      if(formatIndex === '-1') {
        formatIndex = this._cellFormatArray.length.toString();
        this._cellFormatArray.push(formatId);
        const xfElement = this._xmlDocument.createElementNS(this._namespace, 'xf');
        xfElement.setAttribute('numFmtId', formatId.toString());
        this._cellFormatsElement.appendChild(xfElement);
      }
  
      return formatIndex
    }
  
    public toString() {
      const xmlSerializer = new XMLSerializer();
      return xmlSerializer.serializeToString(this._xmlDocument);
    }
}