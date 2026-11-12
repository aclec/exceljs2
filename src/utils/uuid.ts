/**
 * A random RFC 4122 version 4 UUID, using the Web Crypto API (Node >= 19, Bun, browsers). crypto.randomUUID is
 * restricted to secure contexts in browsers, so getRandomValues is the fallback.
 */
export default function uuidv4(): string {
    const { crypto } = globalThis;
    if (typeof crypto.randomUUID === "function") {
        return crypto.randomUUID();
    }
    const b = crypto.getRandomValues(new Uint8Array(16));
    // b has 16 bytes: indexes 6 and 8 exist
    b[6] = (b[6]! & 0x0f) | 0x40;
    b[8] = (b[8]! & 0x3f) | 0x80;
    const h = Array.from(b, (x) => x.toString(16).padStart(2, "0")).join("");
    return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
}
