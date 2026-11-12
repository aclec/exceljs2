/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
import BaseXform from "../base-xform";

class BaseCellAnchorXform extends BaseXform {
    parseOpen(node) {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case this.tag:
                this.reset();
                this.model = {
                    range: {
                        editAs: node.attributes.editAs || "oneCell",
                    },
                };
                break;
            default:
                this.parser = this.map[node.name];
                if (this.parser) {
                    this.parser.parseOpen(node);
                }
                break;
        }
        return true;
    }

    parseText(text) {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    reconcilePicture(model, options) {
        if (model && model.rId) {
            const rel = options.rels[model.rId];
            // take the part after the last '/media/' without a backtracking regex
            // (/.*\/media\/(.+)/ was cubic on Targets like '/media//media//media/...')
            const mediaPos = rel && rel.Target ? rel.Target.lastIndexOf("/media/") : -1;
            const match = mediaPos !== -1 && rel.Target.slice(mediaPos + 7).match(/^(.+[.][a-zA-Z]{3,4})/);
            if (match) {
                const name = match[1];
                const mediaId = options.mediaIndex[name];
                return options.media[mediaId];
            }
        }
        return undefined;
    }
}

export default BaseCellAnchorXform;
