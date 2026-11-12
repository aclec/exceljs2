import type { SaxTag } from "../../utils/parse-sax";
import BaseXform, { type AnyXform } from "./base-xform";

/**
 * Base class of the xforms of an element whose children are parsed by the xforms of {@link BaseXform.map}: default
 * implementations of `parseOpen` / `parseText` / `parseClose`. A subclass provides {@link tag}, the {@link map},
 * {@link createNewModel} and {@link onParserClose}, and its own `render`.
 *
 * Parsing: the start tag of {@link tag} creates the model with {@link createNewModel}; a child element found in `map`
 * is delegated to its xform until it closes, then {@link onParserClose} stores its model; the end tag of {@link tag}
 * completes the parse.
 *
 * @example
 * class DatabarXform extends CompositeXform<DatabarModel> {
 *     declare map: { readonly cfvo: CfvoXform; readonly color: ColorXform };
 *     get tag(): string {
 *         return "dataBar";
 *     }
 *     override createNewModel(node: SaxTag): DatabarModel {
 *         return { cfvo: [], showValue: BaseXform.toBoolValue(node.attributes.showValue) };
 *     }
 *     override onParserClose(name: string): void {
 *         // read the child through its typed reference, not the untyped `parser`
 *         if (name === "cfvo") this.model!.cfvo.push(this.map.cfvo.model!);
 *     }
 * }
 *
 * @typeParam TModel - the model of the element
 */
abstract class CompositeXform<TModel = unknown> extends BaseXform<TModel> {
    /** The name of the element this xform parses. */
    abstract get tag(): string;

    /** The child xforms by element name (each entry an xform, not a wrapper object). */
    declare map?: Readonly<Record<string, AnyXform>>;

    /**
     * Creates the model when the start tag of {@link tag} is read.
     *
     * @param node - the start tag (its attributes usually initialise the model)
     */
    abstract createNewModel(node: SaxTag): TModel;

    /**
     * Opens the element or a child element.
     *
     * @returns `true` when the element is this xform's own or one of its children (found in `map`), else `false`
     */
    override parseOpen(node: SaxTag): boolean {
        // Typical pattern for composite xform
        this.parser = this.parser || this.childXform(node.name);
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }

        if (node.name === this.tag) {
            this.model = this.createNewModel(node);
            return true;
        }

        return false;
    }

    /** Forwards text to the child being parsed. */
    override parseText(text: string): void {
        // Default implementation. Send text to child parser
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /**
     * Called when the child being parsed has closed: stores its model in this one.
     *
     * @param name - the name of the child element (a key of `map`)
     * @param parser - its xform
     */
    abstract onParserClose(name: string, parser: AnyXform): void;

    /** @returns `false` when the element of {@link tag} is closed, else `true` */
    override parseClose(name: string): boolean {
        // Default implementation
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.onParserClose(name, this.parser);
                this.parser = undefined;
            }
            return true;
        }

        return name !== this.tag;
    }

    // the xform of a child element: an own entry of the map only, as element names come from the file (`<constructor>`
    // must not find Object.prototype.constructor)
    private childXform(name: string): AnyXform | undefined {
        const { map } = this;
        return map && Object.hasOwn(map, name) ? map[name] : undefined;
    }
}

export default CompositeXform;
