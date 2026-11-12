const BaseXform = require("./base-xform");
const { runSteps } = require("../../utils/pacer");

class ListXform extends BaseXform {
    constructor(options) {
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

    prepare(model, options) {
        runSteps(this.prepareSteps(model, options));
    }

    // prepare(), render() and reconcile() one item at a time: the caller may pause between items (utils/pacer)
    *prepareSteps(model, options) {
        const { childXform } = this;
        if (model) {
            for (let index = 0; index < model.length; index++) {
                if (index in model) {
                    options.index = index;
                    childXform.prepare(model[index], options);
                    yield;
                }
            }
        }
    }

    render(xmlStream, model) {
        runSteps(this.renderSteps(xmlStream, model));
    }

    *renderSteps(xmlStream, model) {
        if (this.always || (model && model.length)) {
            xmlStream.openNode(this.tag, this.$);
            if (this.count) {
                xmlStream.addAttribute(this.$count, (model && model.length) || 0);
            }

            const { childXform } = this;
            const items = model || [];
            for (let index = 0; index < items.length; index++) {
                if (index in items) {
                    childXform.render(xmlStream, items[index], index);
                    yield;
                }
            }

            xmlStream.closeNode();
        } else if (this.empty) {
            xmlStream.leafNode(this.tag);
        }
    }

    parseOpen(node) {
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

    parseText(text) {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    parseClose(name) {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                this.model.push(this.parser.model);
                this.parser = undefined;

                if (this.maxItems && this.model.length > this.maxItems) {
                    throw new Error(`Max ${this.childXform.tag} count (${this.maxItems}) exceeded`);
                }
            }
            return true;
        }

        return false;
    }

    reconcile(model, options) {
        runSteps(this.reconcileSteps(model, options));
    }

    *reconcileSteps(model, options) {
        if (model) {
            const { childXform } = this;
            for (let index = 0; index < model.length; index++) {
                if (index in model) {
                    childXform.reconcile(model[index], options);
                    yield;
                }
            }
        }
    }
}

module.exports = ListXform;
