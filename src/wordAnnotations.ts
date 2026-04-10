import { Paragraph, TextRun } from "docx";
import { formatCommentDate } from "./utils";
import { ExtendedCommentInfo } from "./types";
import { buildDocumentParagraph, WORD_NAMESPACE_URI } from "./wordXml";

const extractReferenceIds = (paragraph: Element, tagName: 'w:commentReference' | 'w:footnoteReference' | 'w:endnoteReference'): string[] => {
  const refs = paragraph.getElementsByTagName(tagName);
  return Array.from(refs)
    .map(ref => ref.getAttribute('w:id') || '')
    .filter(id => id !== '');
};

export const extractCommentIds = (paragraph: Element): string[] => extractReferenceIds(paragraph, 'w:commentReference');
export const extractFootnoteIds = (paragraph: Element): string[] => extractReferenceIds(paragraph, 'w:footnoteReference');
export const extractEndnoteIds = (paragraph: Element): string[] => extractReferenceIds(paragraph, 'w:endnoteReference');

const buildReferenceParagraph = (
  anchorText: string,
  anchorStyle: string,
  suffixText = ': ',
  options: ConstructorParameters<typeof Paragraph>[0] = {}
): Paragraph => {
  return new Paragraph({
    ...(options as object),
    children: [
      new TextRun({ text: anchorText, style: anchorStyle }),
      new TextRun({ text: suffixText, italics: true })
    ],
  });
};

const buildNoteParagraphs = (
  ids: string[],
  xmlDocument: globalThis.Document,
  elementTagName: 'footnote' | 'endnote',
  anchorPrefix: 'Fn' | 'En',
  anchorStyle: 'FootnoteAnchor' | 'EndnoteAnchor'
): (Paragraph | null)[] => {
  const entries = xmlDocument.getElementsByTagNameNS(WORD_NAMESPACE_URI, elementTagName);

  return ids.flatMap(id => {
    const element = Array.from(entries).find(entry => entry.getAttribute('w:id') === id);
    if (!element) return [null];

    const paragraphs = element.getElementsByTagName('w:p');
    if (!paragraphs || paragraphs.length === 0) return [null];

    const identificationParagraph = new Paragraph({
      children: [
        new TextRun({ text: `[${anchorPrefix} ${id}]`, style: anchorStyle }),
        new TextRun({ text: ': ' })
      ]
    });
    const contentParagraphs = Array.from(paragraphs).map(paragraph => buildDocumentParagraph(paragraph));

    return [identificationParagraph, ...contentParagraphs];
  });
};

export const buildComments = (
  ids: string[],
  xmlComments: globalThis.Document,
  extendedCommentsMap?: Map<string, ExtendedCommentInfo>
): (Paragraph | null)[] => {
  const commentElements = ids.map(id => {
    const comments = xmlComments.getElementsByTagNameNS(WORD_NAMESPACE_URI, 'comment');
    return {
      element: Array.from(comments).find(comment => comment.getAttribute('w:id') === id),
      id
    };
  });

  return commentElements.flatMap(({ element, id }) => {
    if (!element) return [null];

    const paragraphs = element.getElementsByTagName('w:p');
    if (!paragraphs || paragraphs.length === 0) return [null];

    const author = element.getAttribute('w:author') || '';
    const initials = element.getAttribute('w:initials') || '';
    const dateStr = element.getAttribute('w:date') || '';
    const commenterName = initials || author;
    const formattedDate = dateStr ? formatCommentDate(dateStr) : '';

    let extendedInfo: ExtendedCommentInfo | undefined;
    if (extendedCommentsMap && paragraphs.length > 0) {
      const lastParagraph = paragraphs[paragraphs.length - 1];
      const paraId = lastParagraph.getAttribute('w14:paraId') || lastParagraph.getAttribute('paraId');
      if (paraId) {
        extendedInfo = extendedCommentsMap.get(paraId);
      }
    }

    const commentAnchorText = `[Cmt ${id}]`;
    let additionalText = '';
    if (commenterName) {
      additionalText += ` (${commenterName}`;
      if (formattedDate) {
        additionalText += `, ${formattedDate}`;
      }
      additionalText += ')';
    } else if (formattedDate) {
      additionalText += ` (${formattedDate})`;
    }

    if (extendedInfo?.done) {
      additionalText += ' ✓';
    }

    additionalText += ': ';

    const identificationParagraph = buildReferenceParagraph(commentAnchorText, 'CommentAnchor', additionalText, {
      ...(extendedInfo?.paraIdParent && {
        indent: { left: 720 }
      })
    });

    const contentParagraphs = Array.from(paragraphs).map(paragraph => buildDocumentParagraph(paragraph));
    return [identificationParagraph, ...contentParagraphs];
  });
};

export const buildFootnotes = (ids: string[], xmlFootnotes: globalThis.Document): (Paragraph | null)[] => {
  return buildNoteParagraphs(ids, xmlFootnotes, 'footnote', 'Fn', 'FootnoteAnchor');
};

export const buildEndnotes = (ids: string[], xmlEndnotes: globalThis.Document): (Paragraph | null)[] => {
  return buildNoteParagraphs(ids, xmlEndnotes, 'endnote', 'En', 'EndnoteAnchor');
};

export const buildParagraphAnnotations = (
  paragraphElement: Element,
  commentsXml: globalThis.Document | null,
  footnotesXml: globalThis.Document | null,
  endnotesXml: globalThis.Document | null,
  extendedCommentsMap?: Map<string, ExtendedCommentInfo>
): (Paragraph | null)[] => {
  const annotations: (Paragraph | null)[] = [];
  const commentIds = extractCommentIds(paragraphElement);
  const footnoteIds = extractFootnoteIds(paragraphElement);
  const endnoteIds = extractEndnoteIds(paragraphElement);

  if (commentsXml && commentIds.length > 0) {
    annotations.push(...buildComments(commentIds, commentsXml, extendedCommentsMap));
  }

  if (footnotesXml && footnoteIds.length > 0) {
    annotations.push(...buildFootnotes(footnoteIds, footnotesXml));
  }

  if (endnotesXml && endnoteIds.length > 0) {
    annotations.push(...buildEndnotes(endnoteIds, endnotesXml));
  }

  return annotations;
};
