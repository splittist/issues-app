/**
 * Utilities for handling Word document numbering and list formatting.
 */

type NumberFormatter = (num: number) => string;

export type NumberingLevelFormat = {
  lvlText: string;
  numFmt: string;
  start: number;
  lvlRestart?: number;
};

export type NumberingLevelOverride = {
  startOverride?: number;
};

export type NumberingCounterState = number[] | Map<string, number[]>;

type ManualNumberingPattern = {
  pattern: RegExp;
};

const ROMAN_VALUES = [1, 4, 5, 9, 10, 40, 50, 90, 100, 400, 500, 900, 1000];
const LOWER_ROMAN_NUMERALS = ['i', 'iv', 'v', 'ix', 'x', 'xl', 'l', 'xc', 'c', 'cd', 'd', 'cm', 'm'];
const UPPER_ROMAN_NUMERALS = ['I', 'IV', 'V', 'IX', 'X', 'XL', 'L', 'XC', 'C', 'CD', 'D', 'CM', 'M'];

const PERIOD_NUMBERING_PATTERNS: ManualNumberingPattern[] = [
  { pattern: /^(\d+(?:\.\d+)*\.)\s+/ },
  { pattern: /^([a-z]+\.)\s+/ },
  { pattern: /^([A-Z]+\.)\s+/ },
  { pattern: /^([ivxlcdm]+\.)\s+/ },
  { pattern: /^([IVXLCDM]+\.)\s+/ },
  { pattern: /^(\(\d+(?:\.\d+)*\))\s+/ },
  { pattern: /^(\([a-z]+\))\s+/ },
  { pattern: /^(\([A-Z]+\))\s+/ },
  { pattern: /^(\([ivxlcdm]+\))\s+/ },
  { pattern: /^(\([IVXLCDM]+\))\s+/ },
];

const WHITESPACE_DELIMITED_NUMBERING_PATTERNS: ManualNumberingPattern[] = [
  { pattern: /^(\d+(?:\.\d+)*)(?:\t|\s{2,})/ },
  { pattern: /^([a-z]{1,2})(?:\t|\s{2,})/ },
  { pattern: /^([A-Z]{1,2})(?:\t|\s{2,})/ },
  { pattern: /^([ivxlcdm]+)(?:\t|\s{2,})/ },
  { pattern: /^([IVXLCDM]+)(?:\t|\s{2,})/ },
];

const MANUAL_NUMBERING_PATTERNS = [
  ...PERIOD_NUMBERING_PATTERNS,
  ...WHITESPACE_DELIMITED_NUMBERING_PATTERNS,
];

const formatNumberSequence = (counters: number[], formats: NumberingLevelFormat[], currentLevel: number): string[] => {
  return counters.slice(0, currentLevel + 1).map((num, index) => {
    const fmt = formats[index]?.numFmt || 'decimal';
    return formatNumber(num, fmt);
  });
};

const resolveCounters = (counters: NumberingCounterState, key: string, maxLevels: number): number[] => {
  if (Array.isArray(counters)) {
    return counters;
  }

  const existingCounters = counters.get(key);
  if (existingCounters) {
    return existingCounters;
  }

  const initializedCounters = initializeCounters(maxLevels);
  counters.set(key, initializedCounters);
  return initializedCounters;
};

const applyLevelUpdate = (
  counters: number[],
  level: number,
  formats: NumberingLevelFormat[]
): number[] => {
  counters[level]++;

  for (let i = level + 1; i < counters.length; i++) {
    const restartLevel = formats[i]?.lvlRestart;
    if (restartLevel === 0) {
      continue;
    }

    const shouldReset =
      restartLevel === undefined
        ? level <= i - 1
        : restartLevel - 1 <= level;

    if (shouldReset) {
      counters[i] = 0;
    }
  }

  return counters;
};

const renderTrackedNumbering = (
  level: number,
  formats: NumberingLevelFormat[],
  counters: number[],
  levelOverride?: NumberingLevelOverride
): string => {
  const startingValue = levelOverride?.startOverride ?? formats[level]?.start ?? 1;
  if (counters[level] === 0) {
    counters[level] = startingValue - 1;
  }
  applyLevelUpdate(counters, level, formats);
  const formattedNumbers = formatNumberSequence(counters, formats, level);
  const currentFormat = formats[level];

  if (currentFormat?.lvlText) {
    return processLvlText(currentFormat.lvlText, formattedNumbers);
  }

  return formattedNumbers.join('.');
};

const buildAlphaSequence = (num: number, baseCharCode: number): string => {
  let letter = '';

  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(baseCharCode + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }

  return letter;
};

const toRoman = (num: number, numerals: string[]): string => {
  let result = '';

  for (let index = ROMAN_VALUES.length - 1; index >= 0; index--) {
    while (num >= ROMAN_VALUES[index]) {
      result += numerals[index];
      num -= ROMAN_VALUES[index];
    }
  }

  return result;
};

const hasManualNumberingWhitespace = (text: string): boolean => {
  return /^[\t\s]{2,}/.test(text) || /^\t/.test(text);
};

const isRecognizedNumberingToken = (numberingText: string): boolean => {
  return (
    /^\d/.test(numberingText) ||
    /^[a-z]{1,2}\.?$/.test(numberingText) ||
    /^[A-Z]{1,2}\.?$/.test(numberingText) ||
    /^[ivxlcdm]+\.?$/.test(numberingText) ||
    /^[IVXLCDM]+\.?$/.test(numberingText) ||
    /^\(\d/.test(numberingText) ||
    /^\([a-z]{1,2}\)$/.test(numberingText) ||
    /^\([A-Z]{1,2}\)$/.test(numberingText) ||
    /^\([ivxlcdm]+\)$/.test(numberingText) ||
    /^\([IVXLCDM]+\)$/.test(numberingText)
  );
};

const extractParagraphNumbering = (
  paragraphElement: Element
): { ilvl: number; numId: string } | undefined => {
  const numPrElement = paragraphElement.getElementsByTagName('w:numPr')[0];
  if (!numPrElement) return undefined;

  const numId = numPrElement.getElementsByTagName('w:numId')[0]?.getAttribute('w:val');
  const ilvl = numPrElement.getElementsByTagName('w:ilvl')[0]?.getAttribute('w:val');
  if (!numId || !ilvl) return undefined;

  return { ilvl: parseInt(ilvl, 10), numId };
};

/**
 * Converts a number to a lowercase letter representation.
 * @param num - The number to convert.
 * @returns The lowercase letter representation of the number.
 */
export const toLowerLetter = (num: number): string => {
  return buildAlphaSequence(num, 97);
};

/**
 * Converts a number to an uppercase letter representation.
 * @param num - The number to convert.
 * @returns The uppercase letter representation of the number.
 */
export const toUpperLetter = (num: number): string => {
  return buildAlphaSequence(num, 65);
};

/**
 * Converts a number to a lowercase Roman numeral representation.
 * @param num - The number to convert.
 * @returns The lowercase Roman numeral representation of the number.
 */
export const toLowerRoman = (num: number): string => {
  return toRoman(num, LOWER_ROMAN_NUMERALS);
};

/**
 * Converts a number to an uppercase Roman numeral representation.
 * @param num - The number to convert.
 * @returns The uppercase Roman numeral representation of the number.
 */
export const toUpperRoman = (num: number): string => {
  return toRoman(num, UPPER_ROMAN_NUMERALS);
};

/**
 * Converts a number to ordinal representation (1st, 2nd, 3rd, etc.).
 * @param num - The number to convert.
 * @returns The ordinal representation of the number.
 */
export const toOrdinal = (num: number): string => {
  const suffix = ['th', 'st', 'nd', 'rd'];
  const value = num % 100;
  return num + (suffix[(value - 20) % 10] || suffix[value] || suffix[0]);
};

/**
 * Converts a number to cardinal text representation (one, two, three, etc.).
 * @param num - The number to convert.
 * @returns The cardinal text representation of the number.
 */
export const toCardinalText = (num: number): string => {
  const ones = ['', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine'];
  const teens = ['ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
  const tens = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
  const hundreds = ['', 'one hundred', 'two hundred', 'three hundred', 'four hundred', 'five hundred', 'six hundred', 'seven hundred', 'eight hundred', 'nine hundred'];

  if (num === 0) return 'zero';
  if (num < 0) return 'negative ' + toCardinalText(-num);
  if (num >= 1000) return num.toString();

  let result = '';

  if (num >= 100) {
    result += hundreds[Math.floor(num / 100)];
    num %= 100;
    if (num > 0) result += ' ';
  }

  if (num >= 20) {
    result += tens[Math.floor(num / 10)];
    num %= 10;
    if (num > 0) result += '-' + ones[num];
  } else if (num >= 10) {
    result += teens[num - 10];
  } else if (num > 0) {
    result += ones[num];
  }

  return result;
};

/**
 * Converts a number to ordinal text representation (first, second, third, etc.).
 * @param num - The number to convert.
 * @returns The ordinal text representation of the number.
 */
export const toOrdinalText = (num: number): string => {
  const ordinals = ['', 'first', 'second', 'third', 'fourth', 'fifth', 'sixth', 'seventh', 'eighth', 'ninth', 'tenth',
    'eleventh', 'twelfth', 'thirteenth', 'fourteenth', 'fifteenth', 'sixteenth', 'seventeenth', 'eighteenth', 'nineteenth', 'twentieth'];

  if (num <= 20 && num > 0) {
    return ordinals[num];
  }

  const cardinalText = toCardinalText(num);
  if (cardinalText.endsWith('one')) {
    return cardinalText.slice(0, -3) + 'first';
  } else if (cardinalText.endsWith('two')) {
    return cardinalText.slice(0, -3) + 'second';
  } else if (cardinalText.endsWith('three')) {
    return cardinalText.slice(0, -5) + 'third';
  } else if (cardinalText.endsWith('five')) {
    return cardinalText.slice(0, -4) + 'fifth';
  } else if (cardinalText.endsWith('eight')) {
    return cardinalText.slice(0, -5) + 'eighth';
  } else if (cardinalText.endsWith('nine')) {
    return cardinalText.slice(0, -4) + 'ninth';
  } else if (cardinalText.endsWith('twelve')) {
    return cardinalText.slice(0, -6) + 'twelfth';
  } else if (cardinalText.endsWith('y')) {
    return cardinalText.slice(0, -1) + 'ieth';
  }

  return cardinalText + 'th';
};

/**
 * Converts a number to number-in-dash format (- 1 -, - 2 -, etc.).
 * @param num - The number to convert.
 * @returns The number-in-dash representation of the number.
 */
export const toNumberInDash = (num: number): string => {
  return `- ${num} -`;
};

const NUMBER_FORMATTERS: Record<string, NumberFormatter> = {
  bullet: () => '•',
  cardinalText: toCardinalText,
  decimal: (num: number) => num.toString(),
  lowerLetter: toLowerLetter,
  lowerRoman: toLowerRoman,
  numberInDash: toNumberInDash,
  ordinal: toOrdinal,
  ordinalText: toOrdinalText,
  upperLetter: toUpperLetter,
  upperRoman: toUpperRoman,
};

/**
 * Formats a number according to the specified format.
 * @param num - The number to format.
 * @param format - The format to apply.
 * @returns The formatted number as a string.
 */
export const formatNumber = (num: number, format: string): string => {
  const formatter = NUMBER_FORMATTERS[format] || NUMBER_FORMATTERS.decimal;
  return formatter(num);
};

/**
 * Processes a lvlText template by replacing placeholders with formatted numbers.
 * @param template - The lvlText template (e.g., "%1.%2(%3)")
 * @param formattedNumbers - Array of formatted numbers for each level
 * @returns The processed text with placeholders replaced
 */
export const processLvlText = (template: string, formattedNumbers: string[]): string => {
  if (!template) {
    return formattedNumbers.join('.');
  }

  let result = template;
  for (let i = 0; i < formattedNumbers.length; i++) {
    const placeholder = `%${i + 1}`;
    result = result.replace(new RegExp(placeholder, 'g'), formattedNumbers[i]);
  }

  return result;
};

/**
 * Builds maps for numbering from the numbering document.
 * @param numberingDoc - The XML document containing numbering information.
 * @returns An object containing maps for numId to abstractNumId and abstractNumId to format.
 */
export const buildNumberingMaps = (numberingDoc: globalThis.Document) => {
  const numIdToAbstractNumId = new Map<string, string>();
  const abstractNumIdToFormat = new Map<string, NumberingLevelFormat[]>();
  const numIdToLevelOverrides = new Map<string, Map<number, NumberingLevelOverride>>();

  const numElements = numberingDoc.getElementsByTagName('w:num');
  for (const numElement of Array.from(numElements)) {
    const numId = numElement.getAttribute('w:numId');
    const abstractNumId = numElement.getElementsByTagName('w:abstractNumId')[0]?.getAttribute('w:val');
    if (numId && abstractNumId) {
      numIdToAbstractNumId.set(numId, abstractNumId);
    }

    if (numId) {
      const levelOverrides = new Map<number, NumberingLevelOverride>();
      const overrideElements = Array.from(numElement.getElementsByTagName('w:lvlOverride'));
      for (const overrideElement of overrideElements) {
        const ilvlValue = overrideElement.getAttribute('w:ilvl');
        if (!ilvlValue) {
          continue;
        }

        const startOverrideValue = overrideElement
          .getElementsByTagName('w:startOverride')[0]
          ?.getAttribute('w:val');
        levelOverrides.set(parseInt(ilvlValue, 10), {
          startOverride: startOverrideValue ? parseInt(startOverrideValue, 10) : undefined,
        });
      }

      numIdToLevelOverrides.set(numId, levelOverrides);
    }
  }

  const abstractNumElements = numberingDoc.getElementsByTagName('w:abstractNum');
  for (const abstractNumElement of Array.from(abstractNumElements)) {
    const abstractNumId = abstractNumElement.getAttribute('w:abstractNumId');
    const lvlElements = abstractNumElement.getElementsByTagName('w:lvl');
    const formats: NumberingLevelFormat[] = [];

    for (const lvlElement of Array.from(lvlElements)) {
      const ilvlValue = lvlElement.getAttribute('w:ilvl');
      if (!ilvlValue) {
        continue;
      }

      formats[parseInt(ilvlValue, 10)] = {
        numFmt: lvlElement.getElementsByTagName('w:numFmt')[0]?.getAttribute('w:val') || '',
        lvlText: lvlElement.getElementsByTagName('w:lvlText')[0]?.getAttribute('w:val') || '',
        start: parseInt(lvlElement.getElementsByTagName('w:start')[0]?.getAttribute('w:val') || '1', 10),
        lvlRestart: (() => {
          const rawValue = lvlElement.getElementsByTagName('w:lvlRestart')[0]?.getAttribute('w:val');
          return rawValue ? parseInt(rawValue, 10) : undefined;
        })(),
      };
    }

    if (abstractNumId) {
      abstractNumIdToFormat.set(abstractNumId, formats);
    }
  }

  return { numIdToAbstractNumId, abstractNumIdToFormat, numIdToLevelOverrides };
};

/**
 * Interface for style numbering information.
 */
export interface StyleNumberingInfo {
  numId?: string;
  ilvl?: string;
}

/**
 * Interface for style information including hierarchy.
 */
export interface StyleInfo {
  id: string;
  name: string;
  basedOn?: string;
  numbering?: StyleNumberingInfo;
}

/**
 * Builds style hierarchy mappings from the styles document.
 * @param stylesDoc - The XML document containing styles information.
 * @returns Maps for style IDs to style information and style hierarchy resolution.
 */
export const buildStyleMaps = (stylesDoc: globalThis.Document) => {
  const styles = new Map<string, StyleInfo>();

  const styleElements = stylesDoc.getElementsByTagName('w:style');
  for (const styleElement of Array.from(styleElements)) {
    const styleId = styleElement.getAttribute('w:styleId');
    const nameElement = styleElement.getElementsByTagName('w:name')[0];
    const styleName = nameElement?.getAttribute('w:val') || '';
    const basedOnElement = styleElement.getElementsByTagName('w:basedOn')[0];
    const basedOn = basedOnElement?.getAttribute('w:val');

    const pPrElement = styleElement.getElementsByTagName('w:pPr')[0];
    let numbering: StyleNumberingInfo | undefined;

    if (pPrElement) {
      const numPrElement = pPrElement.getElementsByTagName('w:numPr')[0];
      if (numPrElement) {
        const numId = numPrElement.getElementsByTagName('w:numId')[0]?.getAttribute('w:val') || undefined;
        const ilvl = numPrElement.getElementsByTagName('w:ilvl')[0]?.getAttribute('w:val') || undefined;
        numbering = { numId, ilvl };
      }
    }

    if (styleId) {
      styles.set(styleId, {
        id: styleId,
        name: styleName,
        basedOn: basedOn || undefined,
        numbering
      });
    }
  }

  return styles;
};

/**
 * Resolves numbering for a style by walking up the style hierarchy.
 * @param styleId - The style ID to resolve.
 * @param styles - Map of style IDs to style information.
 * @returns The numbering information or undefined if not found.
 */
export const resolveStyleNumbering = (styleId: string, styles: Map<string, StyleInfo>): StyleNumberingInfo | undefined => {
  const visited = new Set<string>();
  let currentStyleId = styleId;

  while (currentStyleId && !visited.has(currentStyleId)) {
    visited.add(currentStyleId);
    const style = styles.get(currentStyleId);

    if (!style) {
      break;
    }

    if (style.numbering) {
      return style.numbering;
    }

    currentStyleId = style.basedOn || '';
  }

  return undefined;
};

/**
 * Initializes counters for numbering levels.
 * @param maxLevels - The maximum number of levels.
 * @returns An array of counters initialized to 0.
 */
export const initializeCounters = (maxLevels: number) => {
  return new Array(maxLevels).fill(0);
};

/**
 * Updates counters for numbering levels.
 * @param counters - The current counters.
 * @param ilvl - The current level.
 * @returns The updated counters.
 */
export const updateCounters = (counters: number[], ilvl: number) => {
  counters[ilvl]++;
  for (let i = ilvl + 1; i < counters.length; i++) {
    counters[i] = 0;
  }
  return counters;
};

/**
 * Tracks numbering for a given paragraph element.
 * @param paragraphElement - The XML element representing the paragraph.
 * @param numIdToAbstractNumId - Map of numId to abstractNumId.
 * @param abstractNumIdToFormat - Map of abstractNumId to format.
 * @param counters - The current counters.
 * @returns The numbering as a string or undefined.
 */
export const trackNumbering = (
  paragraphElement: Element,
  numIdToAbstractNumId: Map<string, string>,
  abstractNumIdToFormat: Map<string, NumberingLevelFormat[]>,
  counters: NumberingCounterState,
  numIdToLevelOverrides?: Map<string, Map<number, NumberingLevelOverride>>
): string | undefined => {
  const paragraphNumbering = extractParagraphNumbering(paragraphElement);
  if (!paragraphNumbering) return undefined;

  const abstractNumId = numIdToAbstractNumId.get(paragraphNumbering.numId);
  if (!abstractNumId) return undefined;

  const formats = abstractNumIdToFormat.get(abstractNumId);
  if (!formats) return undefined;

  const resolvedCounters = resolveCounters(counters, paragraphNumbering.numId, formats.length || 9);
  const levelOverride = numIdToLevelOverrides?.get(paragraphNumbering.numId)?.get(paragraphNumbering.ilvl);
  return renderTrackedNumbering(paragraphNumbering.ilvl, formats, resolvedCounters, levelOverride);
};

/**
 * Extracts the paragraph style ID from a given paragraph element.
 * @param paragraphElement - The XML element representing the paragraph.
 * @returns The style ID or undefined if not found.
 */
export const extractParagraphStyle = (paragraphElement: Element): string | undefined => {
  const pPrElement = paragraphElement.getElementsByTagName('w:pPr')[0];
  if (!pPrElement) return undefined;

  const pStyleElement = pPrElement.getElementsByTagName('w:pStyle')[0];
  if (!pStyleElement) return undefined;

  return pStyleElement.getAttribute('w:val') || undefined;
};

/**
 * Tracks numbering from style hierarchy for a given paragraph element.
 * @param paragraphElement - The XML element representing the paragraph.
 * @param styles - Map of style IDs to style information.
 * @param numIdToAbstractNumId - Map of numId to abstractNumId.
 * @param abstractNumIdToFormat - Map of abstractNumId to format.
 * @param counters - The current counters.
 * @returns The numbering as a string or undefined.
 */
export const trackStyleNumbering = (
  paragraphElement: Element,
  styles: Map<string, StyleInfo>,
  numIdToAbstractNumId: Map<string, string>,
  abstractNumIdToFormat: Map<string, NumberingLevelFormat[]>,
  counters: NumberingCounterState,
  numIdToLevelOverrides?: Map<string, Map<number, NumberingLevelOverride>>
): string | undefined => {
  const styleId = extractParagraphStyle(paragraphElement);
  if (!styleId) return undefined;

  const styleNumbering = resolveStyleNumbering(styleId, styles);
  if (!styleNumbering?.numId) return undefined;

  const abstractNumId = numIdToAbstractNumId.get(styleNumbering.numId);
  if (!abstractNumId) return undefined;

  const formats = abstractNumIdToFormat.get(abstractNumId);
  if (!formats) return undefined;

  const ilvl = parseInt(styleNumbering.ilvl || '0', 10);
  const resolvedCounters = resolveCounters(counters, styleNumbering.numId, formats.length || 9);
  const levelOverride = numIdToLevelOverrides?.get(styleNumbering.numId)?.get(ilvl);
  return renderTrackedNumbering(ilvl, formats, resolvedCounters, levelOverride);
};

/**
 * Extracts the raw text content from a paragraph element.
 * @param paragraphElement - The XML element representing the paragraph.
 * @returns The raw text content of the paragraph.
 */
export const extractParagraphText = (paragraphElement: Element): string => {
  const textRuns = Array.from(paragraphElement.getElementsByTagName('w:r'));
  let fullText = '';

  for (const textRun of textRuns) {
    Array.from(textRun.childNodes).forEach(child => {
      switch (child.nodeName) {
        case 'w:t':
        case 'w:delText':
          fullText += child.textContent || '';
          break;
        case 'w:tab':
          fullText += '\t';
          break;
        default:
          break;
      }
    });
  }

  return fullText;
};

/**
 * Detects manual numbering patterns in paragraph text.
 * Looks for decimal, alpha, or roman numbering followed by tabs/spaces.
 * @param paragraphText - The raw text content of the paragraph.
 * @returns The detected numbering string or undefined if no pattern is found.
 */
export const detectManualNumbering = (paragraphText: string): string | undefined => {
  if (!paragraphText || paragraphText.trim().length === 0) {
    return undefined;
  }

  const trimmedText = paragraphText.trim();

  for (const { pattern } of MANUAL_NUMBERING_PATTERNS) {
    const match = trimmedText.match(pattern);
    if (!match) {
      continue;
    }

    const remainingText = trimmedText.substring(match[0].length).trim();
    if (remainingText.length > 0) {
      return match[1];
    }
  }

  return undefined;
};

/**
 * Validates if a detected numbering pattern is likely to be intentional manual numbering.
 * This helps avoid false positives from text that coincidentally starts with numbers/letters.
 * @param numberingText - The detected numbering text.
 * @param paragraphText - The full paragraph text.
 * @returns True if the numbering is likely intentional, false otherwise.
 */
export const validateManualNumbering = (numberingText: string, paragraphText: string): boolean => {
  if (!numberingText || !paragraphText) {
    return false;
  }

  const trimmedText = paragraphText.trim();
  if (!trimmedText.startsWith(numberingText)) {
    return false;
  }

  const afterNumbering = trimmedText.substring(numberingText.length);
  if (!hasManualNumberingWhitespace(afterNumbering)) {
    return false;
  }

  const content = afterNumbering.replace(/^[\t\s]+/, '');
  if (content.length < 3) {
    return false;
  }

  if (numberingText.length > 10) {
    return false;
  }

  return isRecognizedNumberingToken(numberingText);
};

/**
 * Resolves manual numbering only when both detection and validation succeed.
 * @param paragraphText - The full paragraph text.
 * @returns The validated numbering token or undefined.
 */
export const resolveManualNumbering = (paragraphText: string): string | undefined => {
  const detectedNumbering = detectManualNumbering(paragraphText);
  if (!detectedNumbering) {
    return undefined;
  }

  return validateManualNumbering(detectedNumbering, paragraphText) ? detectedNumbering : undefined;
};
