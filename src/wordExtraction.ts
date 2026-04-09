import JSZip from "jszip";
import { Criteria, ExtractedParagraph, ExtendedCommentInfo, ParagraphSource } from "./types";
import {
  buildNumberingMaps,
  buildStyleMaps,
  extractParagraphStyle,
  extractParagraphText,
  NumberingCounterState,
  NumberingLevelFormat,
  NumberingLevelOverride,
  resolveManualNumbering,
  StyleInfo,
  trackNumbering,
  trackStyleNumbering,
} from "./numberingUtils";
import { buildParagraphAnnotations } from "./wordAnnotations";
import {
  buildDocumentParagraph,
  getIndexedXmlDocuments,
  getOptionalXmlDocument,
  getRequiredXmlDocument,
  paragraphHasTextContent,
  parseExtendedComments,
} from "./wordXml";
import { NumberingBehaviorOptions, resolveCarryForwardPolicy } from "./wordNumberingPolicy";

export type NumberingMaps = {
  numIdToAbstractNumId: Map<string, string>;
  abstractNumIdToFormat: Map<string, NumberingLevelFormat[]>;
  numIdToLevelOverrides: Map<string, Map<number, NumberingLevelOverride>>;
};

type NumberingResolutionOptions = NumberingMaps & {
  styles?: Map<string, StyleInfo>;
  directCounters: NumberingCounterState;
  styleCounters?: NumberingCounterState;
};

const paragraphIsInteresting = (paragraph: Element, criteria: Criteria): boolean => {
  const redline =
    paragraph.getElementsByTagName('w:del').length > 0 ||
    paragraph.getElementsByTagName('w:ins').length > 0 ||
    paragraph.getElementsByTagName('w:moveFrom').length > 0 ||
    paragraph.getElementsByTagName('w:moveTo').length > 0;
  const comments = paragraph.getElementsByTagName('w:commentRangeStart').length > 0;
  const footnotes = paragraph.getElementsByTagName('w:footnoteReference').length > 0;
  const endnotes = paragraph.getElementsByTagName('w:endnoteReference').length > 0;
  const highlight = paragraph.getElementsByTagName('w:highlight').length > 0;

  const textRuns = Array.from(paragraph.getElementsByTagName('w:r')).map(run => {
    const textElement = run.getElementsByTagName('w:t')[0];
    return textElement ? textElement.textContent || '' : '';
  });
  const squareBrackets = textRuns.some(text => text.includes('[') || text.includes(']'));

  return (
    (criteria.redline && redline) ||
    (criteria.comments && comments) ||
    (criteria.footnotes && footnotes) ||
    (criteria.endnotes && endnotes) ||
    (criteria.highlight && highlight) ||
    (criteria.squareBrackets && squareBrackets)
  );
};

const resolveParagraphNumbering = (
  paragraphElement: Element,
  {
    styles,
    numIdToAbstractNumId,
    abstractNumIdToFormat,
    numIdToLevelOverrides,
    directCounters,
    styleCounters = directCounters,
  }: NumberingResolutionOptions
): string | undefined => {
  let numberingInfo = trackNumbering(
    paragraphElement,
    numIdToAbstractNumId,
    abstractNumIdToFormat,
    directCounters,
    numIdToLevelOverrides
  );

  if (!numberingInfo && styles) {
    numberingInfo = trackStyleNumbering(
      paragraphElement,
      styles,
      numIdToAbstractNumId,
      abstractNumIdToFormat,
      styleCounters,
      numIdToLevelOverrides
    );
  }

  if (numberingInfo) {
    return numberingInfo;
  }

  const paragraphText = extractParagraphText(paragraphElement);
  return resolveManualNumbering(paragraphText);
};

const processDocumentParagraphs = (
  document: globalThis.Document,
  criteria: Criteria,
  commentsXml: globalThis.Document | null,
  footnotesXml: globalThis.Document | null,
  endnotesXml: globalThis.Document | null,
  source: ParagraphSource,
  sectionNumber: number,
  numberingMaps?: NumberingMaps,
  styles?: Map<string, StyleInfo>,
  directCounters?: NumberingCounterState,
  styleCounters?: NumberingCounterState,
  extendedCommentsMap?: Map<string, ExtendedCommentInfo>
): ExtractedParagraph[] => {
  const paragraphs = Array.from(document.getElementsByTagName('w:p'));
  const extractedParagraphs: ExtractedParagraph[] = [];

  for (const paragraphElement of paragraphs) {
    if (!paragraphIsInteresting(paragraphElement, criteria) || !paragraphHasTextContent(paragraphElement)) {
      continue;
    }

    const comments = buildParagraphAnnotations(paragraphElement, commentsXml, footnotesXml, endnotesXml, extendedCommentsMap);
    const documentParagraph = buildDocumentParagraph(paragraphElement);
    const styleId = extractParagraphStyle(paragraphElement);
    const numberingInfo = numberingMaps && directCounters
      ? resolveParagraphNumbering(paragraphElement, {
          ...numberingMaps,
          styles,
          directCounters,
          styleCounters,
        })
      : undefined;

    extractedParagraphs.push({
      paragraph: documentParagraph,
      comments,
      section: numberingInfo ? undefined : sectionNumber,
      numbering: numberingInfo,
      style: styleId || undefined,
      source,
    });
  }

  return extractedParagraphs;
};

export const extractParagraphs = async (
  file: File,
  criteria: Criteria,
  options?: NumberingBehaviorOptions
): Promise<ExtractedParagraph[]> => {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const documentXml = await getRequiredXmlDocument(zip, 'word/document.xml');
  const commentsXml = await getOptionalXmlDocument(zip, 'word/comments.xml');
  const commentsExtendedXml = await getOptionalXmlDocument(zip, 'word/commentsExtended.xml');
  const headerDocs = await getIndexedXmlDocuments(zip, 'header');
  const footerDocs = await getIndexedXmlDocuments(zip, 'footer');
  const footnotesXml = await getOptionalXmlDocument(zip, 'word/footnotes.xml');
  const endnotesXml = await getOptionalXmlDocument(zip, 'word/endnotes.xml');
  const numberingXml = await getOptionalXmlDocument(zip, 'word/numbering.xml');
  const stylesXml = await getOptionalXmlDocument(zip, 'word/styles.xml');

  const allParagraphs = Array.from(documentXml.getElementsByTagName('w:p'));
  const numberingMaps = numberingXml
    ? buildNumberingMaps(numberingXml)
    : {
        numIdToAbstractNumId: new Map<string, string>(),
        abstractNumIdToFormat: new Map<string, NumberingLevelFormat[]>(),
        numIdToLevelOverrides: new Map<string, Map<number, NumberingLevelOverride>>(),
      };
  const styles = stylesXml ? buildStyleMaps(stylesXml) : new Map<string, StyleInfo>();
  const extendedCommentsMap = commentsExtendedXml ? parseExtendedComments(commentsExtendedXml) : undefined;
  const carryForwardPolicy = resolveCarryForwardPolicy(options);

  let currentSection = 1;
  let currentPage = 1;
  const counters = new Map<string, number[]>();
  const styleCounters = new Map<string, number[]>();
  let previousNumberingInfo: string | null = null;
  const interestingParagraphs: ExtractedParagraph[] = [];

  for (const paragraphElement of allParagraphs) {
    if (paragraphElement.getElementsByTagName('w:sectPr').length > 0) {
      currentSection++;
      currentPage = 1;
      previousNumberingInfo = null;
    } else if (paragraphElement.getElementsByTagName('w:lastRenderedPageBreak').length > 0) {
      currentPage++;
    }

    let numberingInfo = resolveParagraphNumbering(paragraphElement, {
      ...numberingMaps,
      styles: stylesXml ? styles : undefined,
      directCounters: counters,
      styleCounters,
    });

    const styleId = extractParagraphStyle(paragraphElement);
    const paragraphText = extractParagraphText(paragraphElement);

    if (!numberingInfo && previousNumberingInfo && carryForwardPolicy.enabled) {
      const resetCarryForward = carryForwardPolicy.shouldReset({
        currentPage,
        currentSection,
        paragraphElement,
        paragraphText,
        previousNumberingInfo,
        styleId,
      });

      if (resetCarryForward) {
        previousNumberingInfo = null;
      } else {
        numberingInfo = previousNumberingInfo;
      }
    }

    if (paragraphIsInteresting(paragraphElement, criteria) && paragraphHasTextContent(paragraphElement)) {
      const allComments = buildParagraphAnnotations(paragraphElement, commentsXml, footnotesXml, endnotesXml, extendedCommentsMap);
      const documentParagraph = buildDocumentParagraph(paragraphElement);

      interestingParagraphs.push({
        paragraph: documentParagraph,
        comments: allComments,
        section: numberingInfo ? undefined : currentSection,
        page: numberingInfo ? undefined : currentPage,
        numbering: numberingInfo || undefined,
        style: styleId || undefined,
        source: 'document',
      });
    }

    if (numberingInfo) {
      previousNumberingInfo = numberingInfo;
    }
  }

  let sectionForHeaders = 1;
  const headerFooterCounters = new Map<string, number[]>();

  for (const paragraphElement of allParagraphs) {
    if (paragraphElement.getElementsByTagName('w:sectPr').length === 0) {
      continue;
    }

    for (const headerDoc of headerDocs) {
      interestingParagraphs.push(
        ...processDocumentParagraphs(
          headerDoc,
          criteria,
          commentsXml,
          footnotesXml,
          endnotesXml,
          'header',
          sectionForHeaders,
          numberingMaps,
          styles,
          headerFooterCounters,
          headerFooterCounters,
          extendedCommentsMap
        )
      );
    }

    for (const footerDoc of footerDocs) {
      interestingParagraphs.push(
        ...processDocumentParagraphs(
          footerDoc,
          criteria,
          commentsXml,
          footnotesXml,
          endnotesXml,
          'footer',
          sectionForHeaders,
          numberingMaps,
          styles,
          headerFooterCounters,
          headerFooterCounters,
          extendedCommentsMap
        )
      );
    }

    sectionForHeaders++;
  }

  for (const headerDoc of headerDocs) {
    interestingParagraphs.push(
      ...processDocumentParagraphs(
        headerDoc,
        criteria,
        commentsXml,
        footnotesXml,
        endnotesXml,
        'header',
        sectionForHeaders,
        numberingMaps,
        styles,
        headerFooterCounters,
        headerFooterCounters,
        extendedCommentsMap
      )
    );
  }

  for (const footerDoc of footerDocs) {
    interestingParagraphs.push(
      ...processDocumentParagraphs(
        footerDoc,
        criteria,
        commentsXml,
        footnotesXml,
        endnotesXml,
        'footer',
        sectionForHeaders,
        numberingMaps,
        styles,
        headerFooterCounters,
        headerFooterCounters,
        extendedCommentsMap
      )
    );
  }

  return interestingParagraphs;
};
