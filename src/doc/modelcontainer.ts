// A workbook model without a Workbook (exported as `ModelContainer`): `xlsx` reads a file into `model` and writes `model`.

import XLSX from "../xlsx/xlsx";
import type { WorkbookModel } from "../types";

/** Holds a workbook model, with the xlsx reader / writer working on it (`container.xlsx.readFile(...)`). */
class ModelContainer {
    /** The workbook model: what `xlsx` writes, and what it sets when reading. */
    model: WorkbookModel;
    /** Created on first use. */
    declare private _xlsx?: XLSX;

    constructor(model: WorkbookModel) {
        this.model = model;
    }

    /** The xlsx reader / writer of the model (created on first use). */
    get xlsx(): XLSX {
        if (!this._xlsx) {
            this._xlsx = new XLSX(this);
        }
        return this._xlsx;
    }
}

export default ModelContainer;
