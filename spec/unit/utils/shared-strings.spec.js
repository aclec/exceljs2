const SharedStrings = verquire('utils/shared-strings');

describe('SharedStrings', () => {
  it('Stores and shares string values', () => {
    const ss = new SharedStrings();

    const iHello = ss.add('Hello');
    const iHelloV2 = ss.add('Hello');
    const iGoodbye = ss.add('Goodbye');

    expect(iHello).to.equal(iHelloV2);
    expect(iGoodbye).to.not.equal(iHelloV2);

    expect(ss.count).to.equal(2);
    expect(ss.totalRefs).to.equal(3);
  });

  it('Does not escape values', () => {
    // that's the job of the xml utils
    const ss = new SharedStrings();

    const iXml = ss.add('<tag>value</tag>');
    const iAmpersand = ss.add('&');

    expect(ss.getString(iXml)).to.equal('<tag>value</tag>');
    expect(ss.getString(iAmpersand)).to.equal('&');
  });

  it('Shares rich text values by content', () => {
    const ss = new SharedStrings();
    const bold = {richText: [{font: {bold: true}, text: 'Hello'}]};
    const italic = {richText: [{font: {italic: true}, text: 'Hello'}]};

    const iBold = ss.add(bold);
    const iItalic = ss.add(italic);
    const iBoldV2 = ss.add({richText: [{font: {bold: true}, text: 'Hello'}]});
    // a plain string that looks like a serialised rich text is a different value
    const iJson = ss.add(JSON.stringify(bold.richText));
    const iJsonObject = ss.add(JSON.stringify(bold));
    const iObject = ss.add('[object Object]');

    expect(new Set([iBold, iItalic, iJson, iJsonObject, iObject]).size).to.equal(5);
    expect(iBoldV2).to.equal(iBold);
    expect(ss.getString(iBold)).to.equal(bold);
    expect(ss.getString(iItalic)).to.equal(italic);
    expect(ss.getString(iJson)).to.equal(JSON.stringify(bold.richText));
    expect(ss.count).to.equal(5);
    expect(ss.totalRefs).to.equal(6);
  });
});
