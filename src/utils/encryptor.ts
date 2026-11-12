import * as crypto from "crypto";

/**
 * The hashing of the passwords of sheet protection (ECMA-376 Part 1, 18.3.1.85 sheetProtection: hashValue, saltValue,
 * spinCount). Uses Node's crypto; the browser bundle has a shim for these three functions.
 */
const Encryptor = {
    /**
     * Calculates a hash of the concatenated buffers with the given algorithm.
     * @param algorithm - A Node.js hash algorithm name, e.g. "sha512".
     * @returns The hash.
     */
    hash(algorithm: string, ...buffers: readonly Uint8Array[]): Buffer {
        const hash = crypto.createHash(algorithm);
        hash.update(Buffer.concat(buffers));
        return hash.digest();
    },
    /**
     * Converts a password into its hash: the hash of salt + password (UTF-16LE), then spinCount times the hash of the
     * previous hash + the iteration number (32-bit little endian), as Excel does it (which is why crypto.pbkdf2 cannot).
     * @param password - The password.
     * @param hashAlgorithm - The hash algorithm, e.g. "SHA-512" or "sha512" (case insensitive).
     * @param saltValue - The salt, base64.
     * @param spinCount - The number of iterations.
     * @returns The hash, base64.
     * @throws Error when the hash algorithm is not supported.
     */
    convertPasswordToHash(password: string, hashAlgorithm: string, saltValue: string, spinCount: number): string {
        hashAlgorithm = hashAlgorithm.toLowerCase();
        const hashes = crypto.getHashes();
        if (hashes.indexOf(hashAlgorithm) < 0) {
            throw new Error(`Hash algorithm '${hashAlgorithm}' not supported!`);
        }

        // Password must be in unicode buffer
        const passwordBuffer = Buffer.from(password, "utf16le");
        // Generate the initial hash
        let key = this.hash(hashAlgorithm, Buffer.from(saltValue, "base64"), passwordBuffer);
        // Now regenerate until spin count
        for (let i = 0; i < spinCount; i++) {
            const iterator = Buffer.alloc(4);
            // this is the 'special' element of Excel password hashing
            // that stops us from using crypto.pbkdf2()
            iterator.writeUInt32LE(i, 0);
            key = this.hash(hashAlgorithm, key, iterator);
        }
        return key.toString("base64");
    },
    /**
     * Generates cryptographically strong pseudo-random data.
     * @param size - The number of bytes to generate.
     */
    randomBytes(size: number): Buffer {
        return crypto.randomBytes(size);
    },
};
export default Encryptor;
