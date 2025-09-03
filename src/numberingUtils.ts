/**
 * Utilities for handling Word document numbering and list formatting.
 */

/**
 * Converts a number to a lowercase letter representation.
 * @param num - The number to convert.
 * @returns The lowercase letter representation of the number.
 */
export const toLowerLetter = (num: number): string => {
  let letter = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(97 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
};

/**
 * Converts a number to a lowercase Roman numeral representation.
 * @param num - The number to convert.
 * @returns The lowercase Roman numeral representation of the number.
 */
export const toLowerRoman = (num: number): string => {
  const romanNumerals = ['i', 'iv', 'v', 'ix', 'x', 'xl', 'l', 'xc', 'c', 'cd', 'd', 'cm', 'm'];
  const values = [1, 4, 5, 9, 10, 40, 50, 90, 100, 400, 500, 900, 1000];
  let result = '';
  for (let i = values.length - 1; i >= 0; i--) {
    while (num >= values[i]) {
      result += romanNumerals[i];
      num -= values[i];
    }
  }
  return result;
};

/**
 * Formats a number according to the specified format.
 * @param num - The number to format.
 * @param format - The format to apply.
 * @returns The formatted number as a string.
 */
export const formatNumber = (num: number, format: string): string => {
  switch (format) {
    case 'decimal':
      return num.toString();
    case 'lowerLetter':
      return toLowerLetter(num);
    case 'lowerRoman':
      return toLowerRoman(num);
    default:
      return num.toString();
  }
};

/**
 * Builds maps for numbering from the numbering document.
 * @param numberingDoc - The XML document containing numbering information.
 * @returns An object containing maps for numId to abstractNumId and abstractNumId to format.
 */
export const buildNumberingMaps = (numberingDoc: globalThis.Document) => {
  const numIdToAbstractNumId = new Map<string, string>();
  const abstractNumIdToFormat = new Map<string, { numFmt: string, lvlText: string }[]>();

  const numElements = numberingDoc.getElementsByTagName('w:num');
  for (const numElement of Array.from(numElements)) {
    const numId = numElement.getAttribute('w:numId');
    const abstractNumId = numElement.getElementsByTagName('w:abstractNumId')[0]?.getAttribute('w:val');
    if (numId && abstractNumId) {
      numIdToAbstractNumId.set(numId, abstractNumId);
    }
  }

  const abstractNumElements = numberingDoc.getElementsByTagName('w:abstractNum');
  for (const abstractNumElement of Array.from(abstractNumElements)) {
    const abstractNumId = abstractNumElement.getAttribute('w:abstractNumId');
    const lvlElements = abstractNumElement.getElementsByTagName('w:lvl');
    const formats = Array.from(lvlElements).map(lvlElement => {
      const numFmt = lvlElement.getElementsByTagName('w:numFmt')[0]?.getAttribute('w:val') || '';
      const lvlText = lvlElement.getElementsByTagName('w:lvlText')[0]?.getAttribute('w:val') || '';
      return { numFmt, lvlText };
    });
    if (abstractNumId) {
      abstractNumIdToFormat.set(abstractNumId, formats);
    }
  }

  return { numIdToAbstractNumId, abstractNumIdToFormat };
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
  
  // Parse all styles
  const styleElements = stylesDoc.getElementsByTagName('w:style');
  for (const styleElement of Array.from(styleElements)) {
    const styleId = styleElement.getAttribute('w:styleId');
    const nameElement = styleElement.getElementsByTagName('w:name')[0];
    const styleName = nameElement?.getAttribute('w:val') || '';
    const basedOnElement = styleElement.getElementsByTagName('w:basedOn')[0];
    const basedOn = basedOnElement?.getAttribute('w:val');
    
    // Look for numbering in paragraph properties
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
export const trackNumbering = (paragraphElement: Element, numIdToAbstractNumId: Map<string, string>, abstractNumIdToFormat: Map<string, { numFmt: string, lvlText: string }[]>, counters: number[]): string | undefined => {
  const numPrElement = paragraphElement.getElementsByTagName('w:numPr')[0];
  if (!numPrElement) return undefined;

  const numId = numPrElement.getElementsByTagName('w:numId')[0]?.getAttribute('w:val');
  const ilvl = numPrElement.getElementsByTagName('w:ilvl')[0]?.getAttribute('w:val');
  if (!numId || !ilvl) return undefined;

  const abstractNumId = numIdToAbstractNumId.get(numId);
  if (!abstractNumId) return undefined;

  const formats = abstractNumIdToFormat.get(abstractNumId);
  if (!formats) return undefined;

  // const format = formats[parseInt(ilvl, 10)];
  counters = updateCounters(counters, parseInt(ilvl, 10));
  const numbering = counters.slice(0, parseInt(ilvl, 10) + 1)
      .map((num, index) => {
          const fmt = formats[index]?.numFmt || 'decimal';
          return formatNumber(num, fmt);
       })
    .join('.');
  return numbering;
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
  abstractNumIdToFormat: Map<string, { numFmt: string, lvlText: string }[]>, 
  counters: number[]
): string | undefined => {
  const styleId = extractParagraphStyle(paragraphElement);
  if (!styleId) return undefined;
  
  const styleNumbering = resolveStyleNumbering(styleId, styles);
  if (!styleNumbering || !styleNumbering.numId || !styleNumbering.ilvl) return undefined;
  
  const abstractNumId = numIdToAbstractNumId.get(styleNumbering.numId);
  if (!abstractNumId) return undefined;

  const formats = abstractNumIdToFormat.get(abstractNumId);
  if (!formats) return undefined;

  const ilvl = parseInt(styleNumbering.ilvl, 10);
  counters = updateCounters(counters, ilvl);
  const numbering = counters.slice(0, ilvl + 1)
      .map((num, index) => {
          const fmt = formats[index]?.numFmt || 'decimal';
          return formatNumber(num, fmt);
       })
    .join('.');
  return numbering;
};