import {
  AlignmentType,
  Footer,
  Header,
  ISectionOptions,
  PageNumber,
  PageOrientation,
  Paragraph,
  SectionType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import { dateToday } from "./utils";
import { ExtractedParagraph } from "./types";

export const formatParagraphReference = ({
  section,
  page,
  numbering,
  source,
}: Pick<ExtractedParagraph, 'section' | 'page' | 'numbering' | 'source'>): string => {
  if (numbering) {
    return numbering;
  }

  if (source === 'header' || source === 'footer') {
    return `Sect ${section}, ${source.charAt(0).toUpperCase() + source.slice(1)}`;
  }

  return `Sect ${section}, p ${page}`;
};

export const hasAnyAnnotations = (extractedParagraphs: ExtractedParagraph[][]): boolean => {
  return extractedParagraphs.some(paragraphGroup =>
    paragraphGroup.some(extractedParagraph =>
      extractedParagraph.comments.some(comment => comment !== null)
    )
  );
};

export const buildSections = (extractedParagraphs: ExtractedParagraph[][], names: string[]): ISectionOptions[] => {
  const includeAnnotations = hasAnyAnnotations(extractedParagraphs);

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
            })
          ],
        }),
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.LEFT,
              text: fileName,
            })
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ children: [PageNumber.CURRENT] })],
            })
          ],
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
              children: includeAnnotations ? [
                new TableCell({
                  children: [new Paragraph({ text: "Ref", style: 'Strong' })],
                  width: { size: 8, type: WidthType.PERCENTAGE },
                }),
                new TableCell({
                  children: [new Paragraph({ text: "Paragraph", style: 'Strong' })],
                  width: { size: 31, type: WidthType.PERCENTAGE },
                }),
                new TableCell({
                  children: [new Paragraph({ text: "Annotation", style: 'Strong' })],
                  width: { size: 31, type: WidthType.PERCENTAGE },
                }),
                new TableCell({
                  children: [new Paragraph({ text: "Response", style: 'Strong' })],
                  width: { size: 30, type: WidthType.PERCENTAGE },
                }),
              ] : [
                new TableCell({
                  children: [new Paragraph({ text: "Ref", style: 'Strong' })],
                  width: { size: 8, type: WidthType.PERCENTAGE },
                }),
                new TableCell({
                  children: [new Paragraph({ text: "Paragraph", style: 'Strong' })],
                  width: { size: 46, type: WidthType.PERCENTAGE },
                }),
                new TableCell({
                  children: [new Paragraph({ text: "Response", style: 'Strong' })],
                  width: { size: 46, type: WidthType.PERCENTAGE },
                }),
              ],
            }),
            ...paragraphGroup.map(({ paragraph, comments, section, page, numbering, source }) => {
              const reference = formatParagraphReference({ section, page, numbering, source });

              return new TableRow({
                children: includeAnnotations ? [
                  new TableCell({
                    children: [new Paragraph(reference)],
                  }),
                  new TableCell({
                    children: [paragraph],
                  }),
                  new TableCell({
                    children: comments.map(comment => comment || new Paragraph('')),
                  }),
                  new TableCell({
                    children: [new Paragraph('')],
                  }),
                ] : [
                  new TableCell({
                    children: [new Paragraph(reference)],
                  }),
                  new TableCell({
                    children: [paragraph],
                  }),
                  new TableCell({
                    children: [new Paragraph('')],
                  }),
                ],
              });
            }),
          ],
        }),
      ],
    };
  });
};
