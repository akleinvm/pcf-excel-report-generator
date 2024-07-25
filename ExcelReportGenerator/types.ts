export type PowerAppsObject = {Value: string}
export type PowerAppsArray = {Value: Array<PowerAppsObject>}
export type CellFormat = {Type: string | null, Style: string | null}
export type CellObject = {Value: number, Format: CellFormat}
export type CellIndex = {RowIndex: number, ColumnIndex: number}