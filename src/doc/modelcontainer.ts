/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import XLSX from "../xlsx/xlsx";

class ModelContainer {
    constructor(model) {
        this.model = model;
    }

    get xlsx() {
        if (!this._xlsx) {
            this._xlsx = new XLSX(this);
        }
        return this._xlsx;
    }
}

export default ModelContainer;
