import JSZip from "jszip";
import { Criteria, 
    ExtractedParagraph,
    ParagraphSource,
    } from "./types";
import { Paragraph, 
    ParagraphChild, 
    TextRun, 
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
    } from "docx";
import { dateToday, formatCommentDate } from "./utils";
import { 
    buildNumberingMaps,
    StyleInfo,
    buildStyleMaps,
    initializeCounters,
    trackNumbering,
    extractParagraphStyle,
    trackStyleNumbering
} from "./numberingUtils";
import { buildRunProps } from "./styleUtils";

// Re-export buildStyles for external use
export { buildStyles } from "./styleUtils";

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
        return Array.from(child.getElementsByTagName('w:r')).flatMap(run => buildTextRun(run, "Deletion"));
      case 'w:ins':
        return Array.from(child.getElementsByTagName('w:r')).flatMap(run => buildTextRun(run, "Insertion"));
      case 'w:moveFrom':
        return Array.from(child.getElementsByTagName('w:r')).flatMap(run => buildTextRun(run, "MoveFrom"));
      case 'w:moveTo':
        return Array.from(child.getElementsByTagName('w:r')).flatMap(run => buildTextRun(run, "MoveTo"));
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
    return {
      element: Array.from(comments).find(comment => comment.getAttribute('w:id') === id),
      id: id
    };
  });

  return commentElements.flatMap(({ element, id }) => {
    if (!element) return [null];
    
    const paragraphs = element.getElementsByTagName('w:p');
    if (!paragraphs || paragraphs.length === 0) return [null];
    
    // Extract commenter details
    const author = element.getAttribute('w:author') || '';
    const initials = element.getAttribute('w:initials') || '';
    const dateStr = element.getAttribute('w:date') || '';
    
    // Use initials if available, otherwise use author
    const commenterName = initials || author;
    
    // Format the date if available
    const formattedDate = dateStr ? formatCommentDate(dateStr) : '';
    
    // Build identification text
    let identificationText = `Comment ${id}`;
    if (commenterName) {
      identificationText += ` (${commenterName}`;
      if (formattedDate) {
        identificationText += `, ${formattedDate}`;
      }
      identificationText += ')';
    } else if (formattedDate) {
      identificationText += ` (${formattedDate})`;
    }
    identificationText += ': ';
    
    // Create identification paragraph
    const identificationParagraph = new Paragraph({
      children: [new TextRun({ text: identificationText, italics: true })]
    });
    
    const contentParagraphs = Array.from(paragraphs).map(paragraph => buildDocumentParagraph(paragraph));
    
    return [identificationParagraph, ...contentParagraphs];
  });
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
    return {
      element: Array.from(footnotes).find(footnote => footnote.getAttribute('w:id') === id),
      id: id
    };
  });

  return footnoteElements.flatMap(({ element, id }) => {
    if (!element) return [null];
    
    const paragraphs = element.getElementsByTagName('w:p');
    if (!paragraphs || paragraphs.length === 0) return [null];
    
    // Create identification paragraph
    const identificationParagraph = new Paragraph({
      children: [new TextRun({ text: `Footnote ${id}: ` })]
    });
    
    const contentParagraphs = Array.from(paragraphs).map(paragraph => buildDocumentParagraph(paragraph));
    
    return [identificationParagraph, ...contentParagraphs];
  });
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
    return {
      element: Array.from(endnotes).find(endnote => endnote.getAttribute('w:id') === id),
      id: id
    };
  });

  return endnoteElements.flatMap(({ element, id }) => {
    if (!element) return [null];
    
    const paragraphs = element.getElementsByTagName('w:p');
    if (!paragraphs || paragraphs.length === 0) return [null];
    
    // Create identification paragraph
    const identificationParagraph = new Paragraph({
      children: [new TextRun({ text: `Endnote ${id}: ` })]
    });
    
    const contentParagraphs = Array.from(paragraphs).map(paragraph => buildDocumentParagraph(paragraph));
    
    return [identificationParagraph, ...contentParagraphs];
  });
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
 * Builds one or more TextRun objects from a given run element.
 * @param runElement - The XML element representing the run.
 * @param style - The style to apply.
 * @returns An array of TextRun objects.
 */
const buildTextRun = (runElement: Element, style: string = ''): TextRun[] => {
  const runProps = buildRunProps(runElement.getElementsByTagName('w:rPr')[0], style);
  
  const results: TextRun[] = [];
  let currentText = '';
  
  Array.from(runElement.childNodes).forEach(child => {
    switch(child.nodeName) {
      case 'w:t':
      case 'w:delText':
        currentText += child.textContent || '';
        break;
      case 'w:noBreakHyphen':
        currentText += '‑'; // Non-breaking hyphen character
        break;
      case 'w:softHyphen':
        currentText += '­'; // Soft hyphen character
        break;
      case 'w:cr':
        currentText += '\r';
        break;
      case 'w:br':
        currentText += '\n';
        break;
      case 'w:tab':
        currentText += '\t';
        break;
      case 'w:commentReference': {
        // If we have accumulated text, create a TextRun for it
        if (currentText) {
          results.push(new TextRun({
            text: currentText,
            ...runProps
          }));
          currentText = '';
        }
        
        // Create a styled TextRun for the comment anchor
        const commentId = (child as Element).getAttribute('w:id');
        const commentText = commentId ? `[Comment ${commentId}]` : '[Comment]';
        results.push(new TextRun({
          text: commentText,
          style: 'CommentAnchor',
          ...runProps
        }));
        break;
      }
      case 'w:footnoteReference': {
        // Add any remaining current text as a TextRun first
        if (currentText) {
          results.push(new TextRun({
            text: currentText,
            ...runProps
          }));
          currentText = '';
        }
        
        // Create a styled TextRun for the footnote anchor
        const footnoteId = (child as Element).getAttribute('w:id');
        const footnoteText = footnoteId ? `[Footnote ${footnoteId}]` : '[Footnote]';
        results.push(new TextRun({
          text: footnoteText,
          style: 'FootnoteAnchor',
          ...runProps
        }));
        break;
      }
      case 'w:endnoteReference': {
        // Add any remaining current text as a TextRun first
        if (currentText) {
          results.push(new TextRun({
            text: currentText,
            ...runProps
          }));
          currentText = '';
        }
        
        // Create a styled TextRun for the endnote anchor
        const endnoteId = (child as Element).getAttribute('w:id');
        const endnoteText = endnoteId ? `[Endnote ${endnoteId}]` : '[Endnote]';
        results.push(new TextRun({
          text: endnoteText,
          style: 'EndnoteAnchor',
          ...runProps
        }));
        break;
      }
      default:
        // Ignore unknown elements
        break;
    }
  });
  
  // If we have any remaining text, create a TextRun for it
  if (currentText) {
    results.push(new TextRun({
      text: currentText,
      ...runProps
    }));
  }
  
  // If no results, return a single empty TextRun to maintain compatibility
  return results.length > 0 ? results : [new TextRun({ text: '', ...runProps })];
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

  // Process headers/footers for the final section
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



