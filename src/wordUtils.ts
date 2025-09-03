import JSZip from "jszip";
import { Criteria, 
    ExtractedParagraph,
    IRunContent,
    Break,
    HLColor,
    ULType,
    ParagraphSource,
    } from "./types";
import { Paragraph, 
    ParagraphChild, 
    TextRun, 
    IRunOptions, 
    NoBreakHyphen, 
    SoftHyphen, 
    CarriageReturn, 
    Tab,
    Header,
    Footer,
    AlignmentType,
    SectionType,
    PageNumber,
    Table,
    TableCell, 
    TableRow, 
    WidthType, 
    PageOrientation,
    ISectionOptions, 
    UnderlineType,
    } from "docx";
import { dateToday } from "./utils";

/**
 * Builds a Paragraph object from a given XML element.
 * @param paragraphElement - The XML element representing the paragraph.
 * @returns A Paragraph object.
 */
const buildDocumentParagraph = (paragraphElement: Element): Paragraph => {
  const children = Array.from(paragraphElement.children).flatMap(child => {
    switch(child.nodeName) {
      case 'w:r':
        return buildTextRun(child as Element);
      case 'w:del':
        return Array.from(child.getElementsByTagName('w:r')).map(run => buildTextRun(run, "Deletion"));
      case 'w:ins':
        return Array.from(child.getElementsByTagName('w:r')).map(run => buildTextRun(run, "Insertion"));
      case 'w:moveFrom':
        return Array.from(child.getElementsByTagName('w:r')).map(run => buildTextRun(run, "MoveFrom"));
      case 'w:moveTo':
        return Array.from(child.getElementsByTagName('w:r')).map(run => buildTextRun(run, "MoveTo"));
      default:
        return null;
    }
  }).filter(child => child !== null) as ParagraphChild[];

  return new Paragraph({
    children,
  });
};

/**
 * Extracts comment IDs from a given paragraph element.
 * @param paragraph - The XML element representing the paragraph.
 * @returns An array of comment IDs.
 */
const extractCommentIds = (paragraph: Element):string[] => {
  const commentRefs = paragraph.getElementsByTagName('w:commentReference');
  return Array.from(commentRefs)
    .map(ref => ref.getAttribute('w:id') || '')
    .filter(id => id !== '');
};

/**
 * Extracts footnote IDs from a given paragraph element.
 * @param paragraph - The XML element representing the paragraph.
 * @returns An array of footnote IDs.
 */
const extractFootnoteIds = (paragraph: Element):string[] => {
  const footnoteRefs = paragraph.getElementsByTagName('w:footnoteReference');
  return Array.from(footnoteRefs)
    .map(ref => ref.getAttribute('w:id') || '')
    .filter(id => id !== '');
};

/**
 * Extracts endnote IDs from a given paragraph element.
 * @param paragraph - The XML element representing the paragraph.
 * @returns An array of endnote IDs.
 */
const extractEndnoteIds = (paragraph: Element):string[] => {
  const endnoteRefs = paragraph.getElementsByTagName('w:endnoteReference');
  return Array.from(endnoteRefs)
    .map(ref => ref.getAttribute('w:id') || '')
    .filter(id => id !== '');
};

/**
 * Builds comments from given comment IDs and XML comments document.
 * @param ids - An array of comment IDs.
 * @param xmlComments - The XML document containing comments.
 * @returns An array of Paragraph objects or null.
 */
const buildComments = (ids: string[], xmlComments: globalThis.Document): (Paragraph | null)[] => {
  const namespaceURI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const commentElements = ids.map(id => {
    const comments = xmlComments.getElementsByTagNameNS(namespaceURI,'comment');
    return Array.from(comments).find(comment => comment.getAttribute('w:id') === id);
  });

  return commentElements.flatMap(element => {
    const paragraphs = element?.getElementsByTagName('w:p');
    return paragraphs ? Array.from(paragraphs).map(paragraph => buildDocumentParagraph(paragraph)) : [null];
  })
};

/**
 * Builds footnotes from given footnote IDs and XML footnotes document.
 * @param ids - An array of footnote IDs.
 * @param xmlFootnotes - The XML document containing footnotes.
 * @returns An array of Paragraph objects or null.
 */
const buildFootnotes = (ids: string[], xmlFootnotes: globalThis.Document): (Paragraph | null)[] => {
  const namespaceURI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const footnoteElements = ids.map(id => {
    const footnotes = xmlFootnotes.getElementsByTagNameNS(namespaceURI,'footnote');
    return Array.from(footnotes).find(footnote => footnote.getAttribute('w:id') === id);
  });

  return footnoteElements.flatMap(element => {
    const paragraphs = element?.getElementsByTagName('w:p');
    return paragraphs ? Array.from(paragraphs).map(paragraph => buildDocumentParagraph(paragraph)) : [null];
  })
};

/**
 * Builds endnotes from given endnote IDs and XML endnotes document.
 * @param ids - An array of endnote IDs.
 * @param xmlEndnotes - The XML document containing endnotes.
 * @returns An array of Paragraph objects or null.
 */
const buildEndnotes = (ids: string[], xmlEndnotes: globalThis.Document): (Paragraph | null)[] => {
  const namespaceURI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const endnoteElements = ids.map(id => {
    const endnotes = xmlEndnotes.getElementsByTagNameNS(namespaceURI,'endnote');
    return Array.from(endnotes).find(endnote => endnote.getAttribute('w:id') === id);
  });

  return endnoteElements.flatMap(element => {
    const paragraphs = element?.getElementsByTagName('w:p');
    return paragraphs ? Array.from(paragraphs).map(paragraph => buildDocumentParagraph(paragraph)) : [null];
  })
};

/**
 * Retrieves the main document from the given zip file.
 * @param zip - The JSZip object representing the .docx file.
 * @returns A promise that resolves to the main document.
 */
const getMainDocument = async (zip: JSZip): Promise<globalThis.Document> => {
  const documentXml = await zip.file('word/document.xml')?.async('string');
  if (!documentXml) {
    throw new Error('document.xml not found in the .docx file');
  }

  const docParser = new DOMParser();
  return docParser.parseFromString(documentXml, 'application/xml');
}

/**
 * Retrieves the comments document from the given zip file.
 * @param zip - The JSZip object representing the .docx file.
 * @returns A promise that resolves to the comments document or null.
 */
const getCommentsDocument = async (zip: JSZip): Promise<globalThis.Document | null> => {
  const commentsXml = await zip.file('word/comments.xml')?.async('string');
  if (!commentsXml) {
    return null;
  }

  const commentParser = new DOMParser();
  return commentParser.parseFromString(commentsXml, 'application/xml');
}

/**
 * Retrieves the footnotes document from the given zip file.
 * @param zip - The JSZip object representing the .docx file.
 * @returns A promise that resolves to the footnotes document or null.
 */
const getFootnotesDocument = async (zip: JSZip): Promise<globalThis.Document | null> => {
  const footnotesXml = await zip.file('word/footnotes.xml')?.async('string');
  if (!footnotesXml) {
    return null;
  }

  const footnotesParser = new DOMParser();
  return footnotesParser.parseFromString(footnotesXml, 'application/xml');
}

/**
 * Retrieves the endnotes document from the given zip file.
 * @param zip - The JSZip object representing the .docx file.
 * @returns A promise that resolves to the endnotes document or null.
 */
const getEndnotesDocument = async (zip: JSZip): Promise<globalThis.Document | null> => { 
  const endnotesXml = await zip.file('word/endnotes.xml')?.async('string');
  if (!endnotesXml) {
    return null;
  }

  const endnotesParser = new DOMParser();
  return endnotesParser.parseFromString(endnotesXml, 'application/xml');
}

/**
 * Retrieves header documents from the given zip file.
 * @param zip - The JSZip object representing the .docx file.
 * @returns A promise that resolves to an array of header documents.
 */
const getHeaderDocuments = async (zip: JSZip): Promise<globalThis.Document[]> => {
  const headerDocs: globalThis.Document[] = [];
  const parser = new DOMParser();
  
  // Check for headers (header1.xml, header2.xml, etc.)
  let headerIndex = 1;
  while (true) {
    const headerXml = await zip.file(`word/header${headerIndex}.xml`)?.async('string');
    if (!headerXml) break;
    
    const headerDoc = parser.parseFromString(headerXml, 'application/xml');
    headerDocs.push(headerDoc);
    headerIndex++;
  }
  
  return headerDocs;
}

/**
 * Retrieves footer documents from the given zip file.
 * @returns A promise that resolves to an array of footer documents.
 */
const getFooterDocuments = async (zip: JSZip): Promise<globalThis.Document[]> => {
  const footerDocs: globalThis.Document[] = [];
  const parser = new DOMParser();
  
  // Check for footers (footer1.xml, footer2.xml, etc.)
  let footerIndex = 1;
  while (true) {
    const footerXml = await zip.file(`word/footer${footerIndex}.xml`)?.async('string');
    if (!footerXml) break;
    
    const footerDoc = parser.parseFromString(footerXml, 'application/xml');
    footerDocs.push(footerDoc);
    footerIndex++;
  }
  
  return footerDocs;
}

/**
 * Retrieves the numbering document from the given zip file.
 * @param zip - The JSZip object representing the .docx file.
 * @returns A promise that resolves to the numbering document or null.
 */

const getNumberingDocument = async (zip: JSZip): Promise<globalThis.Document | null> => {
  const numberingXml = await zip.file('word/numbering.xml')?.async('string');
  if (!numberingXml) {
    return null;
  }

  const numberingParser = new DOMParser();
  return numberingParser.parseFromString(numberingXml, 'application/xml');
}

/**
 * Retrieves the styles document from the given zip file.
 * @param zip - The JSZip object representing the .docx file.
 * @returns A promise that resolves to the styles document or null.
 */
const getStylesDocument = async (zip: JSZip): Promise<globalThis.Document | null> => {
  const stylesXml = await zip.file('word/styles.xml')?.async('string');
  if (!stylesXml) {
    return null;
  }

  const stylesParser = new DOMParser();
  return stylesParser.parseFromString(stylesXml, 'application/xml');
}

/**
 * Converts a number to a lowercase letter representation.
 * @param num - The number to convert.
 * @returns The lowercase letter representation of the number.
 */
const toLowerLetter = (num: number): string => {
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
const toLowerRoman = (num: number): string => {
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
const formatNumber = (num: number, format: string): string => {
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
const buildNumberingMaps = (numberingDoc: globalThis.Document) => {
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
interface StyleNumberingInfo {
  numId?: string;
  ilvl?: string;
}

/**
 * Interface for style information including hierarchy.
 */
interface StyleInfo {
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
const buildStyleMaps = (stylesDoc: globalThis.Document) => {
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
const resolveStyleNumbering = (styleId: string, styles: Map<string, StyleInfo>): StyleNumberingInfo | undefined => {
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
const initializeCounters = (maxLevels: number) => {
  return new Array(maxLevels).fill(0);
};

/**
 * Updates counters for numbering levels.
 * @param counters - The current counters.
 * @param ilvl - The current level.
 * @returns The updated counters.
 */
const updateCounters = (counters: number[], ilvl: number) => {
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
const trackNumbering = (paragraphElement: Element, numIdToAbstractNumId: Map<string, string>, abstractNumIdToFormat: Map<string, { numFmt: string, lvlText: string }[]>, counters: number[]): string | undefined => {
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
const extractParagraphStyle = (paragraphElement: Element): string | undefined => {
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
const trackStyleNumbering = (
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

/**
 * Determines if a paragraph is interesting based on the given criteria.
 * @param paragraph - The XML element representing the paragraph.
 * @param criteria - The criteria to check against.
 * @returns True if the paragraph is interesting, false otherwise.
 */
const paragraphIsInteresting = (paragraph: Element, criteria: Criteria): boolean => {
    const redline = paragraph.getElementsByTagName('w:del').length > 0 || paragraph.getElementsByTagName('w:ins').length > 0 || paragraph.getElementsByTagName('w:moveFrom').length > 0 || paragraph.getElementsByTagName('w:moveTo').length > 0;
    const comments = paragraph.getElementsByTagName('w:commentRangeStart').length > 0;
    const footnotes = paragraph.getElementsByTagName('w:footnoteReference').length > 0;
    const endnotes = paragraph.getElementsByTagName('w:endnoteReference').length > 0;
    const highlight = paragraph.getElementsByTagName('w:highlight').length > 0;
    
    const textRuns = Array.from(paragraph.getElementsByTagName('w:r')).map(run => {
      const textElement = run.getElementsByTagName('w:t')[0];
      return textElement ? textElement.textContent || '' : '';});
    const squareBrackets = textRuns.some(text =>
      text.includes('[') || text.includes(']'));
  
    return ((criteria.redline && redline) ||
            (criteria.comments && comments) ||
            (criteria.footnotes && footnotes) ||
            (criteria.endnotes && endnotes) ||
            (criteria.highlight && highlight) ||
            (criteria.squareBrackets && squareBrackets));
    };

/**
 * Builds run properties from a given run properties element.
 * @param runPropsElement - The XML element representing the run properties.
 * @param style - The style to apply.
 * @returns An object containing the run properties.
 */
const buildRunProps = (runPropsElement: Element, style: string = ''): IRunOptions => {
  if (!runPropsElement) {
    return {};
  };
  const bold = runPropsElement.getElementsByTagName('w:b').length > 0;
  const italics = runPropsElement.getElementsByTagName('w:i').length > 0;
  const allCaps = runPropsElement.getElementsByTagName('w:caps').length > 0;
  const smallCaps = runPropsElement.getElementsByTagName('w:smallcaps').length > 0;
  const strike = runPropsElement.getElementsByTagName('w:strike').length > 0;
  const doubleStrike = runPropsElement.getElementsByTagName('w:dstrike').length > 0;
  const highlightElement = runPropsElement.getElementsByTagName('w:highlight')
  const highlight: HLColor | false = highlightElement.length > 0 && highlightElement[0].getAttribute('w:val') as HLColor;
  const underlineElement = runPropsElement.getElementsByTagName('w:u');
  const underlineType = underlineElement.length > 0 ? underlineElement[0].getAttribute('w:val') : undefined;
  const underline: { type?: ULType } | undefined = underlineType ? { type: underlineType as ULType } : undefined;
  const verticalAlignment = runPropsElement.getElementsByTagName('w:vertAlign');
  const superScript = verticalAlignment.length > 0 && verticalAlignment[0].getAttribute('w:val') === 'superscript';
  const subScript = verticalAlignment.length > 0 && verticalAlignment[0].getAttribute('w:val') === 'subscript';

  return {
    bold: bold || undefined,
    italics: italics || undefined,
    allCaps: allCaps || undefined,
    smallCaps: smallCaps || undefined,
    strike: strike || undefined,
    doubleStrike: doubleStrike || undefined,
    highlight: highlight || undefined,
    underline: underline || undefined,
    superScript: superScript || undefined,
    subScript: subScript || undefined,
    style: style === '' ? undefined : style,
  };
}

/**
 * Builds a TextRun object from a given run element.
 * @param runElement - The XML element representing the run.
 * @param style - The style to apply.
 * @returns A TextRun object.
 */
const buildTextRun = (runElement: Element, style: string = ''): TextRun => {
  const runProps = buildRunProps(runElement.getElementsByTagName('w:rPr')[0], style);
  
  const children: (string | IRunContent)[] = Array.from(runElement.childNodes).map(child => {
    switch(child.nodeName) {
      case 'w:t':
      case 'w:delText':
        return child.textContent || '';
      case 'w:noBreakHyphen':
        return new NoBreakHyphen();
      case 'w:softHyphen':
        return new SoftHyphen();
      case 'w:cr':
        return new CarriageReturn();
      case 'w:br':
        return new Break();
      case 'w:tab':
        return new Tab();
      default:
        return null;
    }
  }).filter(child => child !== null) as (string | IRunContent)[];

  return new TextRun({
                children,
                ...runProps
  })
};

/**
 * Processes paragraphs from a document (header or footer) and extracts interesting ones.
 * @param document - The XML document to process.
 * @param criteria - The criteria to filter paragraphs.
 * @param commentsXml - The comments document.
 * @param source - The source type ('header' or 'footer').
 * @param sectionNumber - The section number these paragraphs belong to.
 * @param styles - Map of style IDs to style information.
 * @param numIdToAbstractNumId - Map of numId to abstractNumId.
 * @param abstractNumIdToFormat - Map of abstractNumId to format.
 * @param counters - The current counters for numbering.
 * @returns An array of extracted paragraphs.
 */
const processDocumentParagraphs = (
  document: globalThis.Document,
  criteria: Criteria,
  commentsXml: globalThis.Document | null,
  source: ParagraphSource,
  sectionNumber: number,
  styles?: Map<string, StyleInfo>,
  numIdToAbstractNumId?: Map<string, string>,
  abstractNumIdToFormat?: Map<string, { numFmt: string, lvlText: string }[]>,
  counters?: number[]
): ExtractedParagraph[] => {
  const paragraphs = Array.from(document.getElementsByTagName('w:p'));
  const extractedParagraphs: ExtractedParagraph[] = [];

  for (const paragraphElement of paragraphs) {
    if (paragraphIsInteresting(paragraphElement, criteria)) {
      // Check if the paragraph has any text content before including it
      if (paragraphHasTextContent(paragraphElement)) {
        const commentIds = extractCommentIds(paragraphElement);
        const comments = commentsXml ? buildComments(commentIds, commentsXml) : [];
        const documentParagraph = buildDocumentParagraph(paragraphElement);
        const styleId = extractParagraphStyle(paragraphElement);
        
        // Try to get numbering for header/footer paragraphs
        let numberingInfo: string | undefined;
        if (styles && numIdToAbstractNumId && abstractNumIdToFormat && counters) {
          // Check for direct numbering first
          numberingInfo = trackNumbering(paragraphElement, numIdToAbstractNumId, abstractNumIdToFormat, counters);
          // If no direct numbering, check for style-based numbering
          if (!numberingInfo) {
            numberingInfo = trackStyleNumbering(paragraphElement, styles, numIdToAbstractNumId, abstractNumIdToFormat, counters);
          }
        }

        extractedParagraphs.push({
          paragraph: documentParagraph,
          comments,
          section: numberingInfo ? undefined : sectionNumber,
          numbering: numberingInfo,
          style: styleId || undefined,
          source,
        });
      }
    }
  }

  return extractedParagraphs;
};

/**
 * Extracts paragraphs from a given file based on the specified criteria.
 * @param file - The file to extract paragraphs from.
 * @param criteria - The criteria to filter paragraphs.
 * @returns A promise that resolves to an array of extracted paragraphs.
 */
export const extractParagraphs = async (file: File, criteria: Criteria): Promise<ExtractedParagraph[]> => {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const documentXml = await getMainDocument(zip);
  const commentsXml = await getCommentsDocument(zip);
  const headerDocs = await getHeaderDocuments(zip);
  const footerDocs = await getFooterDocuments(zip);
  const footnotesXml = await getFootnotesDocument(zip);
  const endnotesXml = await getEndnotesDocument(zip);
  const numberingXml = await getNumberingDocument(zip);
  const stylesXml = await getStylesDocument(zip);

  const allParagaphs = Array.from(documentXml.getElementsByTagName('w:p'));

  const { numIdToAbstractNumId, abstractNumIdToFormat } = numberingXml ? buildNumberingMaps(numberingXml) : { numIdToAbstractNumId: new Map<string, string>(), abstractNumIdToFormat: new Map<string, { numFmt: string, lvlText: string }[]>() };
  const styles = stylesXml ? buildStyleMaps(stylesXml) : new Map<string, StyleInfo>();

  let currentSection = 1;
  let currentPage = 1;
  const counters = initializeCounters(9);
  const styleCounters = initializeCounters(9); // Separate counters for style-based numbering
  let previousNumberingInfo: string | null = null;

  const interestingParagraphs: ExtractedParagraph[] = [];

  // Process main document paragraphs

  for (const paragraphElement of allParagaphs) {
    if (paragraphElement.getElementsByTagName('w:sectPr').length > 0) {
      currentSection++;
      currentPage = 1;
      previousNumberingInfo = null;
    } else if (paragraphElement.getElementsByTagName('w:lastRenderedPageBreak').length > 0) {
      currentPage++;
    }

    // Check for direct numbering first (w:numPr elements)
    let numberingInfo = numberingXml ? trackNumbering(paragraphElement, numIdToAbstractNumId, abstractNumIdToFormat, counters) : undefined;
    
    // If no direct numbering, check for style-based numbering
    if (!numberingInfo) {
      numberingInfo = stylesXml ? trackStyleNumbering(paragraphElement, styles, numIdToAbstractNumId, abstractNumIdToFormat, styleCounters) : undefined;
    }

    if (!numberingInfo && previousNumberingInfo) {
      numberingInfo = previousNumberingInfo;
    }

    if (paragraphIsInteresting(paragraphElement, criteria)) {
      // Check if the paragraph has any text content before including it
      if (paragraphHasTextContent(paragraphElement)) {
        const commentIds = extractCommentIds(paragraphElement);
        const footnoteIds = extractFootnoteIds(paragraphElement);
        const endnoteIds = extractEndnoteIds(paragraphElement);
        
        let allComments: (Paragraph | null)[] = [];
        
        // Add comments
        if (commentsXml && commentIds.length > 0) {
          allComments = [...allComments, ...buildComments(commentIds, commentsXml)];
        }
        
        // Add footnotes
        if (footnotesXml && footnoteIds.length > 0) {
          allComments = [...allComments, ...buildFootnotes(footnoteIds, footnotesXml)];
        }
        
        // Add endnotes
        if (endnotesXml && endnoteIds.length > 0) {
          allComments = [...allComments, ...buildEndnotes(endnoteIds, endnotesXml)];
        }
        
        const documentParagraph = buildDocumentParagraph(paragraphElement);
        const styleId = extractParagraphStyle(paragraphElement);

        interestingParagraphs.push({
          paragraph: documentParagraph,
          comments: allComments,
          section: numberingInfo ? undefined : currentSection,
          page: numberingInfo ? undefined : currentPage,
          numbering: numberingInfo ? numberingInfo : undefined,
          style: styleId || undefined,
          source: 'document',
        });
      }
    }

    if (numberingInfo) {
      previousNumberingInfo = numberingInfo;
    }
  }

  // Process header paragraphs for each section
  // Reset section counter for headers/footers processing
  let sectionForHeaders = 1;
  const headerFooterCounters = initializeCounters(9); // Separate counters for headers/footers
  
  for (const paragraphElement of allParagaphs) {
    if (paragraphElement.getElementsByTagName('w:sectPr').length > 0) {
      // Process headers and footers for this section
      for (const headerDoc of headerDocs) {
        const headerParagraphs = processDocumentParagraphs(
          headerDoc, 
          criteria, 
          commentsXml, 
          'header', 
          sectionForHeaders,
          styles,
          numIdToAbstractNumId,
          abstractNumIdToFormat,
          headerFooterCounters
        );
        interestingParagraphs.push(...headerParagraphs);
      }
      
      for (const footerDoc of footerDocs) {
        const footerParagraphs = processDocumentParagraphs(
          footerDoc, 
          criteria, 
          commentsXml, 
          'footer', 
          sectionForHeaders,
          styles,
          numIdToAbstractNumId,
          abstractNumIdToFormat,
          headerFooterCounters
        );
        interestingParagraphs.push(...footerParagraphs);
      }
      
      sectionForHeaders++;
    }
  }

  // Process headers/footers for the final section if no section breaks were found
  if (sectionForHeaders === 1) {
    for (const headerDoc of headerDocs) {
      const headerParagraphs = processDocumentParagraphs(
        headerDoc, 
        criteria, 
        commentsXml, 
        'header', 
        1,
        styles,
        numIdToAbstractNumId,
        abstractNumIdToFormat,
        headerFooterCounters
      );
      interestingParagraphs.push(...headerParagraphs);
    }
    
    for (const footerDoc of footerDocs) {
      const footerParagraphs = processDocumentParagraphs(
        footerDoc, 
        criteria, 
        commentsXml, 
        'footer', 
        1,
        styles,
        numIdToAbstractNumId,
        abstractNumIdToFormat,
        headerFooterCounters
      );
      interestingParagraphs.push(...footerParagraphs);
    }
  }

  return interestingParagraphs;
};

/**
 * Builds sections for the document from the extracted paragraphs.
 * @param extractedParagraphs - An array of arrays of extracted paragraphs.
 * @param names - An array of file names.
 * @returns An array of section options.
 */
export const buildSections = (extractedParagraphs: ExtractedParagraph[][], names: string[]): ISectionOptions[] => {
  return extractedParagraphs.map((paragraphGroup, index) => {
    const fileName = names[index];
    if (!fileName) {
      throw new Error(`File name at index ${index} is undefined.`);
    }
    
    return {
        properties: {
          type: SectionType.CONTINUOUS,
          page: {
            size: {
              orientation: PageOrientation.LANDSCAPE,
            },
          },
          titlePage: true,
        },
        headers: {
          first: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                text: dateToday(),
            })],
          }),
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                text: fileName,
            })],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({children: [ PageNumber.CURRENT ]})],
            })],
          }),
        },
        children: [
            new Paragraph({
              text: fileName,
              style: 'FileName',
          }),
            new Table({
              width: {
                size: 100,
                type: WidthType.PERCENTAGE,
              },
              rows: [
                new TableRow({
                  tableHeader: true,
                  children: [
                    new TableCell({
                      children: [new Paragraph({
                                  text: "Ref",
                                  style: 'Strong',
                    })],
                      width: { size: 8, type:WidthType.PERCENTAGE },
                    }),
                    new TableCell({
                      children: [new Paragraph({
                                  text: "Source",
                                  style: 'Strong',
                    })],
                      width: { size: 7, type:WidthType.PERCENTAGE },
                    }),
                    new TableCell({
                      children: [new Paragraph({
                                  text: "Style",
                                  style: 'Strong',
                    })],
                      width: { size: 10, type:WidthType.PERCENTAGE },
                    }),
                    new TableCell({
                      children: [new Paragraph({
                                  text: "Paragraph",
                                  style: 'Strong',
                    })],
                      width: { size: 35, type:WidthType.PERCENTAGE },
                    }),
                    new TableCell({
                      children: [new Paragraph({
                                  text: "Comment",
                                  style: 'Strong',
                    })],
                      width: { size: 40, type: WidthType.PERCENTAGE },
                    }),
                  ],
                }),
                ...paragraphGroup.map(({ paragraph, comments, section, page, numbering, style, source }) => {
                  return new TableRow({
                    children: [
                      new TableCell({
                        children: [new Paragraph(numbering || `Sect ${section}, p ${page}`)],
                      }),
                      new TableCell({
                        children: [new Paragraph(source.charAt(0).toUpperCase() + source.slice(1))],
                      }),
                      new TableCell({
                        children: [new Paragraph(style || '')],
                      }),
                      new TableCell({
                        children: [paragraph,]
                      }),
                      new TableCell({
                        children: comments.map(comment => comment || new Paragraph('')),
                      }),
                    ],
                  });
                }),
              ],
            }),
          ],
      };
    })
  };

/**
 * Checks if a paragraph element (from XML) contains any text content.
 * @param paragraphElement - The XML element representing the paragraph.
 * @returns True if the paragraph has text content, false if it's empty.
 */
const paragraphHasTextContent = (paragraphElement: Element): boolean => {
  // Get all text runs in the paragraph
  const textRuns = Array.from(paragraphElement.getElementsByTagName('w:r'));
  
  // Check each text run for text content
  for (const textRun of textRuns) {
    const textElements = Array.from(textRun.getElementsByTagName('w:t'));
    const delTextElements = Array.from(textRun.getElementsByTagName('w:delText'));
    
    // Check regular text elements
    for (const textElement of textElements) {
      const text = textElement.textContent || '';
      if (text.trim().length > 0) {
        return true;
      }
    }
    
    // Check deleted text elements
    for (const delTextElement of delTextElements) {
      const text = delTextElement.textContent || '';
      if (text.trim().length > 0) {
        return true;
      }
    }
  }
  
  return false;
};

/**
 * Builds styles for the document.
 * @returns An object containing paragraph and character styles.
 */
export const buildStyles = () => {
  return {
    paragraphStyles: [
      {
        id: 'Normal',
        name: 'Normal',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: {
          size: 24,
          font: 'Calibri',
        },
      },
      {
        id: 'FileName',
        name: 'FileName',
        basedOn: 'Normal',
        next: 'Normal',
        quickFormat: true,
        run: {
          bold: true,
        },
        paragraph: {
          spacing: {
            before: 160,
            after: 120,
          },
        }
      },
    ],
    characterStyles: [
      {
        id: 'Deletion',
        name: 'Deletion',
        basedOn: 'Normal',
        run: {
          color: 'FF0000',
          strike: true,
        },
      },
      {
        id: 'Insertion',
        name: 'Insertion',
        basedOn: 'Normal',
        run: {
          color: '0000FF',
          underline: {
            type: UnderlineType.SINGLE,
          },
        },
      },
      {
        id: 'MoveFrom',
        name: 'MoveFrom',
        basedOn: 'Normal',
        run: {
          color: '006400',
          strike: true,
        },
      },
      {
        id: 'MoveTo',
        name: 'MoveTo',
        basedOn: 'Normal',
        run: {
          color: '006400',
          underline: {
            type: UnderlineType.DOUBLE,
          },
        },
      }
    ],
  }
}

