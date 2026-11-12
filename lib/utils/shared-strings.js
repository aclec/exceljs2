class SharedStrings {
  constructor() {
    this._values = [];
    this._totalRefs = 0;
    this._hash = Object.create(null);
    // rich text values, keyed by the JSON of their runs. Kept apart from _hash: as an object
    // key every rich text would be "[object Object]", and a JSON key there could collide
    // with an identical plain string.
    this._richTextHash = new Map();
  }

  get count() {
    return this._values.length;
  }

  get values() {
    return this._values;
  }

  get totalRefs() {
    return this._totalRefs;
  }

  getString(index) {
    return this._values[index];
  }

  add(value) {
    let index;
    if (value && value.richText) {
      const key = JSON.stringify(value.richText);
      index = this._richTextHash.get(key);
      if (index === undefined) {
        index = this._values.length;
        this._richTextHash.set(key, index);
        this._values.push(value);
      }
    } else {
      index = this._hash[value];
      if (index === undefined) {
        index = this._hash[value] = this._values.length;
        this._values.push(value);
      }
    }
    this._totalRefs++;
    return index;
  }
}

module.exports = SharedStrings;
