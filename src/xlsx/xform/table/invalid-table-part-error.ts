// The error of a table part that CT_Table forbids: the document reader skips such a part.

/**
 * A table part that CT_Table forbids (a table or a column without name, no `<tableColumns>`): Excel repairs such a file
 * by removing the table, and the document reader skips the part (`XLSX._processTableEntry`).
 */
class InvalidTablePartError extends Error {
    override readonly name = "InvalidTablePartError";
}

export default InvalidTablePartError;
