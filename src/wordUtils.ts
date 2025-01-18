import JSZip from "jszip";
import { Criteria, 
    ExtractedParagraph,
    IRunContent,
    Break,
    HLColor,
    ULType, 
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

const buildCommentHeadline = (commentElement: Element): Paragraph => {
  const id = commentElement.getAttribute('w:id') || '?';
  const person = commentElement.getAttribute('w:initials') || commentElement.getAttribute('w:author') || '';
  return new Paragraph({
    children: [
      new TextRun({text: `Comment ${id}`, italics: true}),
      new TextRun({text: ' by '}),
      new TextRun({text: person, bold: true}),
      new TextRun({text: ':'}),
    ],
  });
}

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
    if (!element) return [null];
    const commentHeadline = buildCommentHeadline(element);
    const paragraphElements = element.getElementsByTagName('w:p');
    const paragraphs = paragraphElements ? Array.from(paragraphElements).map(paragraph => buildDocumentParagraph(paragraph)) : [null];
    return [commentHeadline, ...paragraphs];
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

/*
const getFootnotesDocument = async (zip: JSZip): Promise<globalThis.Document> => {
  const footnotesXml = await zip.file('word/footnotes.xml')?.async('string');
  if (!footnotesXml) {
    throw new Error('footnotes.xml not found in the .docx file')
  }

  const footnotesParser = new DOMParser();
  return footnotesParser.parseFromString(footnotesXml, 'application/xml');
}

const getEndnotesDocument = async (zip: JSZip): Promise<globalThis.Document> => { 
  const endnotesXml = await zip.file('word/endnotes.xml')?.async('string');
  if (!endnotesXml) {
    throw new Error('endnotes.xml not found in the .docx file')
  }

  const endnotesParser = new DOMParser();
  return endnotesParser.parseFromString(endnotesXml, 'application/xml');
}
*/

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
};

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
const formatNumber = (num: number, format: string | null): string => {
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
 * Builds a map from style IDs to abstractNumIds.
 * @param stylesDoc - The XML document containing styles information.
 * @returns A map from style IDs to abstractNumIds.
 */
const buildStyleToAbstractNumIdMap = (stylesDoc: globalThis.Document): Map<string, string> => {
  const styleToAbstractNumId = new Map<string, string>();
  const styleElements = stylesDoc.getElementsByTagName('w:style');

  for (const styleElement of Array.from(styleElements)) {
    const styleId = styleElement.getAttribute('w:styleId');
    const numIdElement = styleElement.getElementsByTagName('w:numId')[0];
    if (styleId && numIdElement) {
      const numId = numIdElement.getAttribute('w:val');
      if (numId) {
        styleToAbstractNumId.set(styleId, numId);
      }
    }
  }

  return styleToAbstractNumId;
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
 * @param styleToAbstractNumId - Map of style IDs to abstractNumIds.
 * @param counters - The current counters.
 * @returns The numbering as a string or null.
 */
const trackNumbering = (paragraphElement: Element, 
                        numIdToAbstractNumId: Map<string, string>, 
                        abstractNumIdToFormat: Map<string, { numFmt: string, lvlText: string }[]>, 
                        styleToAbstractNumId: Map<string, string>,
                        counters: number[]) => {
  const numPrElement = paragraphElement.getElementsByTagName('w:numPr')[0];
  let numId: string | null = null;
  let ilvl: string | null = null;

  if (numPrElement) {
    numId = numPrElement.getElementsByTagName('w:numId')[0]?.getAttribute('w:val');
    ilvl = numPrElement.getElementsByTagName('w:ilvl')[0]?.getAttribute('w:val');
  }

  if (!numId) {
    const pStyleElement = paragraphElement.getElementsByTagName('w:pStyle')[0];
    const styleId = pStyleElement?.getAttribute('w:val');
    if (styleId) {
      numId = styleToAbstractNumId.get(styleId || '') || null;
      ilvl = '0'; // Default to level 0 if using style-based numbering
    }
  }

  if (!numId || !ilvl) return null;

  const abstractNumId = numIdToAbstractNumId.get(numId);
  if (!abstractNumId) return null;

  const formats = abstractNumIdToFormat.get(abstractNumId);
  if (!formats) return null;

  counters = updateCounters(counters, parseInt(ilvl, 10));
  const numbering = counters.slice(0, parseInt(ilvl, 10) + 1).map((num, index) => formatNumber(num, formats[index]?.numFmt)).join('.');
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
    const highlight = paragraph.getElementsByTagName('w:highlight').length > 0;
    
    const textRuns = Array.from(paragraph.getElementsByTagName('w:r')).map(run => {
      const textElement = run.getElementsByTagName('w:t')[0];
      return textElement ? textElement.textContent || '' : '';});
    const squareBrackets = textRuns.some(text =>
      text.includes('[') || text.includes(']'));
  
    return ((criteria.redline && redline) ||
            (criteria.comments && comments) ||
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
      case 'w:commentReference':
        return `[Comment ${(child as Element).getAttribute('w:id')}]`;
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
  const numberingXml = await getNumberingDocument(zip);
  const stylesXml = await getStylesDocument(zip);

  const allParagaphs = Array.from(documentXml.getElementsByTagName('w:p'));

  const { numIdToAbstractNumId, abstractNumIdToFormat } = numberingXml ? buildNumberingMaps(numberingXml) : { numIdToAbstractNumId: new Map<string, string>(), abstractNumIdToFormat: new Map<string, { numFmt: string, lvlText: string }[]>() };
  const styleToAbstractNumId = stylesXml ? buildStyleToAbstractNumIdMap(stylesXml) : new Map<string, string>();

  let currentSection = 1;
  let currentPage = 1;
  const counters = initializeCounters(9);
  let previousNumberingInfo: string | null = null;

  const interestingParagraphs: ExtractedParagraph[] = [];

  for (const paragraphElement of allParagaphs) {
    if (paragraphElement.getElementsByTagName('w:sectPr').length > 0) {
      currentSection++;
      currentPage = 1;
      previousNumberingInfo = null;
    } else if (paragraphElement.getElementsByTagName('w:lastRenderedPageBreak').length > 0) {
      currentPage++;
    }

    let numberingInfo = numberingXml ? trackNumbering(paragraphElement, numIdToAbstractNumId, abstractNumIdToFormat, styleToAbstractNumId, counters) : undefined;

    if (!numberingInfo && previousNumberingInfo) {
      numberingInfo = previousNumberingInfo;
    }

    if (paragraphIsInteresting(paragraphElement, criteria)) {
      const commentIds = extractCommentIds(paragraphElement);
      const comments = commentsXml ? buildComments(commentIds, commentsXml) : [];
      const documentParagraph = buildDocumentParagraph(paragraphElement);

      interestingParagraphs.push({
        paragraph: documentParagraph,
        comments,
        section: numberingInfo ? undefined : currentSection,
        page: numberingInfo ? undefined : currentPage,
        numbering: numberingInfo ? numberingInfo : undefined,
      });
    }

    if (numberingInfo) {
      previousNumberingInfo = numberingInfo;
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
                      width: { size: 10, type:WidthType.PERCENTAGE },
                    }),
                    new TableCell({
                      children: [new Paragraph({
                                  text: "Paragraph",
                                  style: 'Strong',
                    })],
                      width: { size: 45, type:WidthType.PERCENTAGE },
                    }),
                    new TableCell({
                      children: [new Paragraph({
                                  text: "Comment",
                                  style: 'Strong',
                    })],
                      width: { size: 45, type: WidthType.PERCENTAGE },
                    }),
                  ],
                }),
                ...paragraphGroup.map(({ paragraph, comments, section, page, numbering }) => {
                  return new TableRow({
                    children: [
                      new TableCell({
                        children: [new Paragraph(numbering || `Sect ${section}, p ${page}`)],
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
