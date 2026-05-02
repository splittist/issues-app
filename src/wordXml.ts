import JSZip from "jszip";
import { Paragraph, ParagraphChild, TextRun } from "docx";
import { buildRunProps } from "./styleUtils";
import { ExtendedCommentInfo } from "./types";

export const WORD_NAMESPACE_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const parseXmlDocument = (xml: string): globalThis.Document => {
  const parser = new DOMParser();
  return parser.parseFromString(xml, 'application/xml');
};

export const getRequiredXmlDocument = async (zip: JSZip, path: string): Promise<globalThis.Document> => {
  const xml = await zip.file(path)?.async('string');
  if (!xml) {
    throw new Error(`${path.split('/').pop()} not found in the .docx file`);
  }

  return parseXmlDocument(xml);
};

export const getOptionalXmlDocument = async (zip: JSZip, path: string): Promise<globalThis.Document | null> => {
  const xml = await zip.file(path)?.async('string');
  return xml ? parseXmlDocument(xml) : null;
};

export const getIndexedXmlDocuments = async (zip: JSZip, baseName: 'header' | 'footer'): Promise<globalThis.Document[]> => {
  const documents: globalThis.Document[] = [];
  let index = 1;

  while (true) {
    const xml = await zip.file(`word/${baseName}${index}.xml`)?.async('string');
    if (!xml) {
      break;
    }

    documents.push(parseXmlDocument(xml));
    index++;
  }

  return documents;
};

export const parseExtendedComments = (xmlCommentsExtended: globalThis.Document): Map<string, ExtendedCommentInfo> => {
  const extendedCommentsMap = new Map<string, ExtendedCommentInfo>();
  const commentExElements = xmlCommentsExtended.querySelectorAll('commentEx, w15\\:commentEx');

  for (const commentExElement of commentExElements) {
    const paraId = commentExElement.getAttribute('w15:paraId') || commentExElement.getAttribute('paraId');
    if (paraId) {
      const paraIdParent = commentExElement.getAttribute('w15:paraIdParent') || commentExElement.getAttribute('paraIdParent');
      const doneAttr = commentExElement.getAttribute('w15:done') || commentExElement.getAttribute('done');
      const done = doneAttr === '1';

      const extendedInfo: ExtendedCommentInfo = {
        paraId,
        ...(paraIdParent && { paraIdParent }),
        ...(doneAttr && { done })
      };

      extendedCommentsMap.set(paraId, extendedInfo);
    }
  }

  return extendedCommentsMap;
};

export const buildTextRun = (runElement: Element, style = ''): TextRun[] => {
  const runProps = buildRunProps(runElement.getElementsByTagName('w:rPr')[0], style);
  const results: TextRun[] = [];
  let currentText = '';

  Array.from(runElement.childNodes).forEach(child => {
    switch (child.nodeName) {
      case 'w:t':
      case 'w:delText':
        currentText += child.textContent || '';
        break;
      case 'w:noBreakHyphen':
        currentText += '‑';
        break;
      case 'w:softHyphen':
        currentText += '­';
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
        if (currentText) {
          results.push(new TextRun({ text: currentText, ...runProps }));
          currentText = '';
        }

        const commentId = (child as Element).getAttribute('w:id');
        const commentText = commentId ? `[Cmt ${commentId}]` : '[Cmt]';
        results.push(new TextRun({
          text: commentText,
          ...runProps,
          style: 'CommentAnchor'
        }));
        break;
      }
      case 'w:footnoteReference': {
        if (currentText) {
          results.push(new TextRun({ text: currentText, ...runProps }));
          currentText = '';
        }

        const footnoteId = (child as Element).getAttribute('w:id');
        const footnoteText = footnoteId ? `[Fn ${footnoteId}]` : '[Fn]';
        results.push(new TextRun({
          text: footnoteText,
          ...runProps,
          style: 'FootnoteAnchor'
        }));
        break;
      }
      case 'w:endnoteReference': {
        if (currentText) {
          results.push(new TextRun({ text: currentText, ...runProps }));
          currentText = '';
        }

        const endnoteId = (child as Element).getAttribute('w:id');
        const endnoteText = endnoteId ? `[En ${endnoteId}]` : '[En]';
        results.push(new TextRun({
          text: endnoteText,
          ...runProps,
          style: 'EndnoteAnchor'
        }));
        break;
      }
      default:
        break;
    }
  });

  if (currentText) {
    results.push(new TextRun({ text: currentText, ...runProps }));
  }

  return results.length > 0 ? results : [new TextRun({ text: '', ...runProps })];
};

type RevisionStyle = '' | 'Deletion' | 'Insertion' | 'MoveFrom' | 'MoveTo';

const getRevisionStyle = (nodeName: string, currentStyle: RevisionStyle): RevisionStyle => {
  switch (nodeName) {
    case 'w:del':
      return 'Deletion';
    case 'w:ins':
      return 'Insertion';
    case 'w:moveFrom':
      return 'MoveFrom';
    case 'w:moveTo':
      return 'MoveTo';
    default:
      return currentStyle;
  }
};

const buildParagraphChildren = (element: Element, activeStyle: RevisionStyle = ''): ParagraphChild[] => {
  if (element.nodeName === 'w:r') {
    return buildTextRun(element, activeStyle);
  }

  const nextStyle = getRevisionStyle(element.nodeName, activeStyle);

  return Array.from(element.children).flatMap(child =>
    buildParagraphChildren(child as Element, nextStyle)
  );
};

export const buildDocumentParagraph = (paragraphElement: Element): Paragraph => {
  return new Paragraph({ children: buildParagraphChildren(paragraphElement) });
};

export const paragraphHasTextContent = (paragraphElement: Element): boolean => {
  const textRuns = Array.from(paragraphElement.getElementsByTagName('w:r'));

  for (const textRun of textRuns) {
    const textElements = Array.from(textRun.getElementsByTagName('w:t'));
    const delTextElements = Array.from(textRun.getElementsByTagName('w:delText'));

    for (const textElement of textElements) {
      const text = textElement.textContent || '';
      if (text.trim().length > 0) {
        return true;
      }
    }

    for (const delTextElement of delTextElements) {
      const text = delTextElement.textContent || '';
      if (text.trim().length > 0) {
        return true;
      }
    }
  }

  return false;
};
