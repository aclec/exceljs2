import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { CustomFilterModel, FilterColumnModel, FilterModel } from "../../../types";
import { parseBoolean } from "../../../utils/utils";
import BaseXform, { ownEntry } from "../base-xform";
import ListXform from "../list-xform";

import CustomFilterXform from "./custom-filter-xform";
import FilterXform from "./filter-xform";

/**
 * A `<filterColumn colId hiddenButton>` of an auto filter (see {@link FilterColumnModel}), with its custom filters. The
 * other filters (dynamicFilter, top10...) are skipped when parsing; `<filters>` is parsed but not kept.
 */
class FilterColumnXform extends BaseXform<FilterColumnModel> {
    declare map: { readonly customFilters: ListXform<CustomFilterModel>; readonly filters: ListXform<FilterModel> };

    /** While parsing: the depth inside an element that is skipped (0: none). */
    private skipDepth = 0;

    constructor() {
        super();

        this.map = {
            customFilters: new ListXform({
                tag: "customFilters",
                count: false,
                empty: true,
                childXform: new CustomFilterXform(),
            }),
            filters: new ListXform({
                tag: "filters",
                count: false,
                empty: true,
                childXform: new FilterXform(),
            }),
        };
    }

    get tag(): string {
        return "filterColumn";
    }

    /** Sets `colId`: the index of the column. */
    override prepare(model: FilterColumnModel, options: { readonly index: number }): void {
        model.colId = options.index.toString();
    }

    override render(xmlStream: XmlStream, model: FilterColumnModel): void {
        if (model.customFilters) {
            xmlStream.openNode(this.tag, {
                colId: model.colId,
                hiddenButton: model.filterButton ? "0" : "1",
            });

            this.map.customFilters.render(xmlStream, model.customFilters);

            xmlStream.closeNode();
            return;
        }
        xmlStream.leafNode(this.tag, {
            colId: model.colId,
            hiddenButton: model.filterButton ? "0" : "1",
        });
    }

    /** @returns `true`: every element inside `<filterColumn>` is accepted (the unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        if (this.skipDepth) {
            this.skipDepth++;
            return true;
        }
        const { attributes } = node;
        switch (node.name) {
            case this.tag:
                this.reset();
                this.skipDepth = 0;
                this.model = {
                    colId: attributes.colId,
                    // CT_FilterColumn: hiddenButton defaults to false, i.e. the button is shown
                    filterButton: !parseBoolean(attributes.hiddenButton),
                };
                return true;
            default: {
                const parser = ownEntry(this.map, node.name);
                this.parser = parser;
                if (parser) {
                    parser.parseOpen(node);
                    return true;
                }
                // dynamicFilter, top10, colorFilter, iconFilter, extLst... are not modelled: skip the subtree
                this.skipDepth = 1;
                return true;
            }
        }
    }

    override parseText(): void {}

    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.parser = undefined;
            }
            return true;
        }
        if (this.skipDepth) {
            this.skipDepth--;
            return true;
        }
        switch (name) {
            case this.tag:
                // the model was created when the element opened
                if (this.map.customFilters.model) {
                    this.model!.customFilters = this.map.customFilters.model;
                }
                return false;
            default:
                // could be some unrecognised tags
                return true;
        }
    }
}

export default FilterColumnXform;
