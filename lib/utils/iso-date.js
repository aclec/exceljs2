// Parser for the ISO 8601 values of t="d" cells, shared by the document and the streaming reader.
//
// Accepted: YYYY-MM-DD, optionally followed by Thh:mm, Thh:mm:ss or Thh:mm:ss.fff (any number of
// fraction digits, truncated to milliseconds) and an offset (Z, +hh:mm or -hh:mm, at most 14:00 as
// in XML Schema dateTime).
// - A value without an offset is read as UTC, like serial dates (utils.excelToDate).
// - T24:00 (also 24:00:00 and 24:00:00.000) is the end of the day, i.e. 00:00 of the next day, as
//   ISO 8601 and XML Schema allow it. Any other hour above 23 is invalid.
// - Years 0000-0099 are kept as such (no mapping to 1900-1999).
// Returns a Date, or undefined when the text is not such a value (callers keep it as a string).
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?(Z|([+-])(\d{2}):(\d{2}))?)?$/;

function parseIsoDate(text) {
    const match = typeof text === "string" && ISO_DATE.exec(text);
    if (!match) {
        return undefined;
    }
    const [year, month, day, hours, minutes, seconds] = match.slice(1, 7).map((value) => Number(value || 0));
    const fraction = match[7] || "";
    const ms = Number(`${fraction}000`.slice(0, 3));

    const endOfDay = hours === 24 && minutes === 0 && seconds === 0 && !/[1-9]/.test(fraction);
    if ((hours > 23 && !endOfDay) || minutes > 59 || seconds > 59) {
        return undefined;
    }

    // setUTCFullYear, unlike Date.UTC, does not map years 0-99 to 1900-1999
    const date = new Date(0);
    date.setUTCFullYear(year, month - 1, day);
    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
        // out of range month or day (2024-13-01, 2024-02-30)
        return undefined;
    }
    date.setUTCHours(hours, minutes, seconds, ms);

    if (match[9]) {
        const offsetHours = Number(match[10]);
        const offsetMinutes = Number(match[11]);
        if (offsetMinutes > 59 || offsetHours * 60 + offsetMinutes > 14 * 60) {
            return undefined;
        }
        const offset = (offsetHours * 60 + offsetMinutes) * (match[9] === "-" ? -1 : 1);
        date.setTime(date.getTime() - offset * 60000);
    }
    return date;
}

module.exports = { parseIsoDate };
