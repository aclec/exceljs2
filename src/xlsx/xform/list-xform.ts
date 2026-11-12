import type { SaxTag } from "../../utils/parse-sax";
import type XmlStream from "../../utils/xml-stream";
import type { XmlAttributes } from "../../utils/xml-stream";
import BaseXform from "./base-xform";

/** The options of a {@link ListXform}. */
export interface ListXformOptions<TItem> {
    /** The list element, e.g. `cols`. */
    readonly tag: string;
    /** The xform of one item (one child element); its `tag`, if any, names the items in the `maxItems` error. */
    readonly childXform: ItemXform<TItem>;
    /** Render the list element even when the list is empty or missing. */
    readonly always?: boolean;
    /** Write the number of items as an attribute (named by `$count`). */
    readonly count?: boolean;
    /** Render an empty list element (`<tag/>`) when the list is empty or missing (and `always` is not set). */
    readonly empty?: boolean;
    /** The name of the count attribute. Default: `count`. */
    readonly $count?: string;
    /** Attributes of the list element. */
    readonly $?: XmlAttributes;
    /** Parsing fails when the list has more items than this (a limit against hostile files). Default: no limit. */
    readonly maxItems?: number;
}

/**
 * The options passed by {@link ListXform.prepare} to the item xforms: those of the caller, with `index` set to the
 * index of the item (0-based). The object is the caller's own, mutated.
 */
export interface ListItemOptions {
    index?: number;
}

/** The xform of the items of a {@link ListXform}. */
export type ItemXform<TItem> = BaseXform<TItem> & { readonly tag?: string };

/** The state of a ListXform item step: an item has been processed (the caller may pause). */
type Steps = Generator<void, void, undefined>;

// runs all the steps of a generator at once. pacer.ts has the same runSteps, but it is not typed yet (its static
// Pacer.runSteps is not visible to TypeScript): use its named export once it has one.
function runSteps(steps: Steps): void {
    for (let step = steps.next(); !step.done; step = steps.next());
}

/**
 * The xform of a list element whose children are all items of one kind, e.g. `<cols><col/><col/></cols>`: the model is
 * the array of the item models.
 *
 * `prepare`, `render` and `reconcile` process the items one at a time and have generator versions (`prepareSteps`,
 * `renderSteps`, `reconcileSteps`) that yield after each item, so that a caller can pause between them (utils/pacer).
 * Holes of a sparse array are skipped.
 *
 * @typeParam TItem - the model of one item
 */
class ListXform<TItem> extends BaseXform<TItem[]> {
    /** The list element. */
    readonly tag: string;
    /** See {@link ListXformOptions.always}. */
    readonly always: boolean;
    /** See {@link ListXformOptions.count}. */
    readonly count: boolean | undefined;
    /** See {@link ListXformOptions.empty}. */
    readonly empty: boolean | undefined;
    /** See {@link ListXformOptions.$count}. */
    readonly $count: string;
    /** See {@link ListXformOptions.$}. */
    readonly $: XmlAttributes | undefined;
    /** The xform of one item. */
    readonly childXform: ItemXform<TItem>;
    /** See {@link ListXformOptions.maxItems}. */
    readonly maxItems: number | undefined;

    constructor(options: ListXformOptions<TItem>) {
        super();

        this.tag = options.tag;
        this.always = !!options.always;
        this.count = options.count;
        this.empty = options.empty;
        this.$count = options.$count || "count";
        this.$ = options.$;
        this.childXform = options.childXform;
        this.maxItems = options.maxItems;
    }

    /** Prepares each item with the item xform; `options.index` is set to the index of the item. */
    override prepare(model: readonly TItem[] | undefined, options: ListItemOptions): void {
        runSteps(this.prepareSteps(model, options));
    }

    /** {@link prepare}, one item per step. */
    *prepareSteps(model: readonly TItem[] | undefined, options: ListItemOptions): Steps {
        const { childXform } = this;
        if (model) {
            for (let index = 0; index < model.length; index++) {
                // `in`: holes of a sparse array are skipped (an item that is present may still be undefined)
                if (index in model) {
                    options.index = index;
                    childXform.prepare(model[index] as TItem, options);
                    yield;
                }
            }
        }
    }

    /** Writes the list element and its items (the item xform gets the index of the item as third argument). */
    override render(xmlStream: XmlStream, model: readonly TItem[] | undefined): void {
        runSteps(this.renderSteps(xmlStream, model));
    }

    /** {@link render}, one item per step. */
    *renderSteps(xmlStream: XmlStream, model: readonly TItem[] | undefined): Steps {
        if (this.always || (model && model.length)) {
            xmlStream.openNode(this.tag, this.$);
            if (this.count) {
                xmlStream.addAttribute(this.$count, (model && model.length) || 0);
            }

            const { childXform } = this;
            const items = model || [];
            for (let index = 0; index < items.length; index++) {
                if (index in items) {
                    childXform.render(xmlStream, items[index] as TItem, index);
                    yield;
                }
            }

            xmlStream.closeNode();
        } else if (this.empty) {
            xmlStream.leafNode(this.tag);
        }
    }

    /** @returns `true` for the list element and for an element the item xform accepts */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case this.tag:
                this.model = [];
                return true;
            default:
                if (this.childXform.parseOpen(node)) {
                    this.parser = this.childXform;
                    return true;
                }
                return false;
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /**
     * @returns `false` at the end of the list element
     * @throws Error when the list has more than `maxItems` items
     */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // the list element was opened before its items (this.model = []); the item xform has its model once its
                // element is closed
                this.model!.push(this.childXform.model as TItem);
                this.parser = undefined;

                if (this.maxItems && this.model!.length > this.maxItems) {
                    throw new Error(`Max ${String(this.childXform.tag)} count (${this.maxItems}) exceeded`);
                }
            }
            return true;
        }

        return false;
    }

    /** Reconciles each item with the item xform. */
    override reconcile(model: readonly TItem[] | undefined, options?: unknown): void {
        runSteps(this.reconcileSteps(model, options));
    }

    /** {@link reconcile}, one item per step. */
    *reconcileSteps(model: readonly TItem[] | undefined, options?: unknown): Steps {
        if (model) {
            const { childXform } = this;
            for (let index = 0; index < model.length; index++) {
                if (index in model) {
                    childXform.reconcile(model[index] as TItem, options);
                    yield;
                }
            }
        }
    }
}

export default ListXform;
