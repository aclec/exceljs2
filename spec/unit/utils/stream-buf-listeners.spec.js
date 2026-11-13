const StreamBuf = verquire("utils/stream-buf");

// StreamBuf extends Duplex but emits its events itself: Readable's flow control must not act on it
describe("StreamBuf listeners", () => {
    it("keeps its data when a 'readable' listener is added", async () => {
        const stream = new StreamBuf();
        stream.write("Hello");
        stream.on("readable", () => {});
        await new Promise((resolve) => setImmediate(resolve));
        expect(stream.toBuffer().toString()).to.equal("Hello");
    });

    it("stays paused when a 'data' listener is added", () => {
        const stream = new StreamBuf();
        const chunks = [];
        stream.pause();
        stream.on("data", (chunk) => chunks.push(chunk));
        stream.write("Hello");
        expect(stream.isPaused()).to.equal(true);
        expect(chunks).to.have.length(0);
    });

    it("emits 'data' to its listeners when not paused", () => {
        const stream = new StreamBuf();
        const chunks = [];
        stream.on("data", (chunk) => chunks.push(chunk.toString()));
        stream.write("Hello");
        expect(chunks).to.deep.equal(["Hello"]);
    });

    it("does not support async iteration (it would never end)", () => {
        const stream = new StreamBuf();
        expect(() => stream[Symbol.asyncIterator]()).to.throw("Not Implemented");
    });
});
