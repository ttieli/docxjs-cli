/**
 * Normalize style configuration to ensure docx receives numeric values.
 * Converts common measurement strings (cm/in/mm/pt/twip) to twips.
 */
function parseMeasurement(val) {
    if (val === null || val === undefined) return undefined;
    if (typeof val === 'number') return val;

    if (typeof val === 'string') {
        const trimmed = val.trim();
        const num = parseFloat(trimmed);
        if (isNaN(num)) return undefined;

        if (/cm$/i.test(trimmed)) return Math.round(num * 566.9291339); // 1cm = 566.929... twips
        if (/mm$/i.test(trimmed)) return Math.round(num * 56.69291339);  // 1mm = 56.69 twips
        if (/in(ch)?$/i.test(trimmed)) return Math.round(num * 1440);    // 1in = 1440 twips
        if (/pt$/i.test(trimmed)) return Math.round(num * 20);           // 1pt = 20 twips
        if (/twips?$/i.test(trimmed)) return Math.round(num);

        // Fallback: plain number string assumed to already be twips
        return Math.round(num);
    }
    return undefined;
}

function toNumber(val) {
    if (val === null || val === undefined) return undefined;
    const num = Number(val);
    return Number.isFinite(num) ? num : undefined;
}

function normalizeStyleConfig(style = {}) {
    const normalized = { ...style };

    // Normalize margin
    if (normalized.margin) {
        const margin = { ...normalized.margin };
        ['top', 'bottom', 'left', 'right'].forEach(side => {
            const parsed = parseMeasurement(margin[side]);
            if (parsed !== undefined) margin[side] = parsed;
            else delete margin[side];
        });
        normalized.margin = margin;
    }

    // Normalize numeric fields
    const numericFields = [
        'fontSizeMain', 'fontSizeH1', 'fontSizeH2', 'fontSizeH3',
        'fontSizeH4', 'fontSizeH5', 'fontSizeH6',
        'lineSpacing', 'paragraphIndent'
    ];
    numericFields.forEach(key => {
        if (normalized[key] !== undefined) {
            const num = toNumber(normalized[key]);
            if (num !== undefined) normalized[key] = num;
            else delete normalized[key];
        }
    });

    // Table numeric fields
    if (normalized.table) {
        const table = { ...normalized.table };
        ['borderSize', 'width'].forEach(key => {
            if (table[key] !== undefined) {
                const num = toNumber(table[key]);
                if (num !== undefined) table[key] = num;
                else delete table[key];
            }
        });
        normalized.table = table;
    }

    return normalized;
}

module.exports = { normalizeStyleConfig };
