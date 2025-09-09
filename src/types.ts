import { XmlComponent,
    Paragraph, 
    HighlightColor, 
    UnderlineType,
    NoBreakHyphen, 
    SoftHyphen, 
    CarriageReturn, 
    Tab,
    } from "docx";

// This is a workaround for the fact that the docx library does not export Break as a named export
export class Break extends XmlComponent {
  constructor() {
    super("w:br");
  }
}

/**
 * Interface representing the criteria for extracting paragraphs.
 */
export interface Criteria {
  redline: boolean;
  highlight: boolean;
  squareBrackets: boolean;
  comments: boolean;
  footnotes: boolean;
  endnotes: boolean;
}

/**
 * Type representing the source of a paragraph.
 */
export type ParagraphSource = 'document' | 'header' | 'footer';

/**
 * Interface representing an extracted paragraph.
 */
export interface ExtractedParagraph {
  paragraph: Paragraph;
  comments: (Paragraph | null)[];
  section?: number;
  page?: number;
  numbering?: string;
  style?: string;
  source: ParagraphSource;
}

/**
 * Type representing highlight colors.
 */
export type HLColor = (typeof HighlightColor)[keyof typeof HighlightColor] | null;

/**
 * Type representing underline types.
 */
export type ULType = (typeof UnderlineType)[keyof typeof UnderlineType];

/**
 * Type representing run content.
 */
export type IRunContent = string | NoBreakHyphen | SoftHyphen | CarriageReturn | Break | Tab;

/**
 * Interface representing extended comment information from commentsExtended.xml.
 */
export interface ExtendedCommentInfo {
  paraId: string;
  paraIdParent?: string;
  done?: boolean;
}

/**
 * Interface representing a mapping of comment IDs to their extended information.
 */
export interface ExtendedCommentsMap {
  [commentId: string]: ExtendedCommentInfo;
}
