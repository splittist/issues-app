/**
 * Utilities for handling Word document styles and text run formatting.
 */

import { 
    IRunOptions, 
    UnderlineType,
    ShadingType
} from "docx";
import { HLColor, ULType } from "./types";

/**
 * Builds run properties from a given run properties element.
 * @param runPropsElement - The XML element representing the run properties.
 * @param style - The style to apply.
 * @returns An object containing the run properties.
 */
export const buildRunProps = (runPropsElement: Element, style: string = ''): IRunOptions => {
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
      },
      {
        id: 'CommentAnchor',
        name: 'CommentAnchor',
        basedOn: 'Normal',
        run: {
          shading: {
            type: ShadingType.PERCENT_50,
            fill: 'C0C0C0', // Light gray
          },
        },
      },
      {
        id: 'FootnoteAnchor',
        name: 'FootnoteAnchor',
        basedOn: 'Normal',
        run: {
          shading: {
            type: ShadingType.PERCENT_50,
            fill: 'ADD8E6', // Light blue
          },
        },
      },
      {
        id: 'EndnoteAnchor',
        name: 'EndnoteAnchor',
        basedOn: 'Normal',
        run: {
          shading: {
            type: ShadingType.PERCENT_50,
            fill: '90EE90', // Light green
          },
        },
      }
    ],
  }
}