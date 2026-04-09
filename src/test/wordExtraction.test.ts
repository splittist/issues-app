import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { extractParagraphs } from '../wordExtraction'
import { Criteria } from '../types'

const createDocxFile = async (files: Record<string, string>, name = 'test.docx'): Promise<File> => {
  const zip = new JSZip()

  Object.entries(files).forEach(([path, content]) => {
    zip.file(path, content)
  })

  const data = await zip.generateAsync({ type: 'uint8array' })
  const file = new File([data], name)
  Object.defineProperty(file, 'arrayBuffer', {
    value: async () => data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength),
  })
  return file
}

const baseCriteria: Criteria = {
  redline: false,
  highlight: false,
  squareBrackets: false,
  comments: false,
  footnotes: false,
  endnotes: false,
}

const createCommentedParagraph = ({
  commentId,
  text,
  numbering,
  styleId,
}: {
  commentId: string
  text: string
  numbering?: { numId: string; ilvl: string }
  styleId?: string
}): string => `
  <w:p>
    <w:pPr>
      ${styleId ? `<w:pStyle w:val="${styleId}" />` : ''}
      ${numbering ? `
        <w:numPr>
          <w:ilvl w:val="${numbering.ilvl}" />
          <w:numId w:val="${numbering.numId}" />
        </w:numPr>
      ` : ''}
    </w:pPr>
    <w:commentRangeStart w:id="${commentId}" />
    <w:r><w:t>${text}</w:t></w:r>
    <w:r><w:commentReference w:id="${commentId}" /></w:r>
  </w:p>
`

const createCommentsXml = (commentIds: string[]): string => `
  <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    ${commentIds.map(commentId => `
      <w:comment w:id="${commentId}" w:author="Test Author" w:initials="TA" w:date="2025-01-15T00:00:00Z">
        <w:p><w:r><w:t>Comment ${commentId}</w:t></w:r></w:p>
      </w:comment>
    `).join('')}
  </w:comments>
`

describe('wordExtraction', () => {
  it('should extract paragraphs containing insertions and deletions when redline is enabled', async () => {
    const file = await createDocxFile({
      'word/document.xml': `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r><w:t>Normal paragraph</w:t></w:r>
            </w:p>
            <w:p>
              <w:r><w:t>Before </w:t></w:r>
              <w:del><w:r><w:delText>old</w:delText></w:r></w:del>
              <w:ins><w:r><w:t>new</w:t></w:r></w:ins>
            </w:p>
          </w:body>
        </w:document>
      `,
    })

    const extracted = await extractParagraphs(file, { ...baseCriteria, redline: true })

    expect(extracted).toHaveLength(1)
    expect(extracted[0]?.source).toBe('document')
  })

  it('should include comment annotations for interesting paragraphs', async () => {
    const file = await createDocxFile({
      'word/document.xml': `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:commentRangeStart w:id="0" />
              <w:r><w:t>Commented paragraph</w:t></w:r>
              <w:r><w:commentReference w:id="0" /></w:r>
            </w:p>
          </w:body>
        </w:document>
      `,
      'word/comments.xml': `
        <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:comment w:id="0" w:author="Jane Doe" w:initials="JD" w:date="2025-01-15T00:00:00Z">
            <w:p>
              <w:r><w:t>Review note</w:t></w:r>
            </w:p>
          </w:comment>
        </w:comments>
      `,
    })

    const extracted = await extractParagraphs(file, { ...baseCriteria, comments: true })

    expect(extracted).toHaveLength(1)
    expect(extracted[0]?.comments.length).toBeGreaterThan(0)
  })

  it('should carry forward the last rendered numbering string to a later interesting paragraph', async () => {
    const file = await createDocxFile({
      'word/document.xml': `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr>
                <w:numPr>
                  <w:ilvl w:val="0" />
                  <w:numId w:val="7" />
                </w:numPr>
              </w:pPr>
              <w:r><w:t>1. Definitions</w:t></w:r>
            </w:p>
            ${createCommentedParagraph({
              commentId: '0',
              text: 'The definition applies to the rest of this clause.',
            })}
          </w:body>
        </w:document>
      `,
      'word/comments.xml': createCommentsXml(['0']),
      'word/numbering.xml': `
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="10">
            <w:lvl w:ilvl="0">
              <w:numFmt w:val="decimal" />
              <w:lvlText w:val="%1." />
            </w:lvl>
          </w:abstractNum>
          <w:num w:numId="7">
            <w:abstractNumId w:val="10" />
          </w:num>
        </w:numbering>
      `,
    })

    const extracted = await extractParagraphs(file, { ...baseCriteria, comments: true })

    expect(extracted).toHaveLength(1)
    expect(extracted[0]?.numbering).toBe('1.')
    expect(extracted[0]?.section).toBeUndefined()
    expect(extracted[0]?.page).toBeUndefined()
  })

  it('should resolve style-linked numbering for interesting paragraphs', async () => {
    const file = await createDocxFile({
      'word/document.xml': `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr>
                <w:pStyle w:val="Heading1" />
              </w:pPr>
              <w:r><w:t>Services</w:t></w:r>
            </w:p>
            ${createCommentedParagraph({
              commentId: '0',
              text: 'Scope of Services',
              styleId: 'Heading2',
            })}
          </w:body>
        </w:document>
      `,
      'word/comments.xml': createCommentsXml(['0']),
      'word/styles.xml': `
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="paragraph" w:styleId="Heading1">
            <w:name w:val="Heading 1" />
            <w:pPr>
              <w:numPr>
                <w:numId w:val="5" />
              </w:numPr>
            </w:pPr>
          </w:style>
          <w:style w:type="paragraph" w:styleId="Heading2">
            <w:name w:val="Heading 2" />
            <w:basedOn w:val="Heading1" />
            <w:pPr>
              <w:numPr>
                <w:ilvl w:val="1" />
                <w:numId w:val="5" />
              </w:numPr>
            </w:pPr>
          </w:style>
        </w:styles>
      `,
      'word/numbering.xml': `
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="3">
            <w:lvl w:ilvl="0">
              <w:numFmt w:val="decimal" />
              <w:lvlText w:val="%1." />
            </w:lvl>
            <w:lvl w:ilvl="1">
              <w:numFmt w:val="decimal" />
              <w:lvlText w:val="%1.%2" />
            </w:lvl>
          </w:abstractNum>
          <w:num w:numId="5">
            <w:abstractNumId w:val="3" />
          </w:num>
        </w:numbering>
      `,
    })

    const extracted = await extractParagraphs(file, { ...baseCriteria, comments: true })

    expect(extracted).toHaveLength(1)
    expect(extracted[0]?.numbering).toBe('1.1')
    expect(extracted[0]?.style).toBe('Heading2')
  })

  it('should fall back to section and page references when numbering metadata is absent', async () => {
    const file = await createDocxFile({
      'word/document.xml': `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            ${createCommentedParagraph({
              commentId: '0',
              text: 'A commented paragraph without numbering metadata.',
            })}
          </w:body>
        </w:document>
      `,
      'word/comments.xml': createCommentsXml(['0']),
    })

    const extracted = await extractParagraphs(file, { ...baseCriteria, comments: true })

    expect(extracted).toHaveLength(1)
    expect(extracted[0]?.numbering).toBeUndefined()
    expect(extracted[0]?.section).toBe(1)
    expect(extracted[0]?.page).toBe(1)
  })

  it('should restart numbering when a new numId starts a fresh list instance', async () => {
    const file = await createDocxFile({
      'word/document.xml': `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr>
                <w:numPr>
                  <w:ilvl w:val="0" />
                  <w:numId w:val="1" />
                </w:numPr>
              </w:pPr>
              <w:r><w:t>First list item</w:t></w:r>
            </w:p>
            <w:p>
              <w:pPr>
                <w:numPr>
                  <w:ilvl w:val="0" />
                  <w:numId w:val="1" />
                </w:numPr>
              </w:pPr>
              <w:r><w:t>Second list item</w:t></w:r>
            </w:p>
            ${createCommentedParagraph({
              commentId: '0',
              text: 'A restarted list should render as item 1 again.',
              numbering: { numId: '2', ilvl: '0' },
            })}
          </w:body>
        </w:document>
      `,
      'word/comments.xml': createCommentsXml(['0']),
      'word/numbering.xml': `
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="10">
            <w:lvl w:ilvl="0">
              <w:numFmt w:val="decimal" />
              <w:lvlText w:val="%1." />
            </w:lvl>
          </w:abstractNum>
          <w:num w:numId="1">
            <w:abstractNumId w:val="10" />
          </w:num>
          <w:num w:numId="2">
            <w:abstractNumId w:val="10" />
          </w:num>
        </w:numbering>
      `,
    })

    const extracted = await extractParagraphs(file, { ...baseCriteria, comments: true })

    expect(extracted).toHaveLength(1)
    expect(extracted[0]?.numbering).toBe('1.')
  })

  it('should honor startOverride when numbering.xml defines a restarted sequence', async () => {
    const file = await createDocxFile({
      'word/document.xml': `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            ${createCommentedParagraph({
              commentId: '0',
              text: 'This clause should start at 4.',
              numbering: { numId: '9', ilvl: '0' },
            })}
          </w:body>
        </w:document>
      `,
      'word/comments.xml': createCommentsXml(['0']),
      'word/numbering.xml': `
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="12">
            <w:lvl w:ilvl="0">
              <w:start w:val="1" />
              <w:numFmt w:val="decimal" />
              <w:lvlText w:val="%1." />
            </w:lvl>
          </w:abstractNum>
          <w:num w:numId="9">
            <w:abstractNumId w:val="12" />
            <w:lvlOverride w:ilvl="0">
              <w:startOverride w:val="4" />
            </w:lvlOverride>
          </w:num>
        </w:numbering>
      `,
    })

    const extracted = await extractParagraphs(file, { ...baseCriteria, comments: true })

    expect(extracted).toHaveLength(1)
    expect(extracted[0]?.numbering).toBe('4.')
  })

  it('should honor lvlRestart=0 by preserving lower-level numbering across higher-level paragraphs', async () => {
    const file = await createDocxFile({
      'word/document.xml': `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr>
                <w:numPr>
                  <w:ilvl w:val="0" />
                  <w:numId w:val="21" />
                </w:numPr>
              </w:pPr>
              <w:r><w:t>Top level one</w:t></w:r>
            </w:p>
            <w:p>
              <w:pPr>
                <w:numPr>
                  <w:ilvl w:val="1" />
                  <w:numId w:val="21" />
                </w:numPr>
              </w:pPr>
              <w:r><w:t>Second level a</w:t></w:r>
            </w:p>
            <w:p>
              <w:pPr>
                <w:numPr>
                  <w:ilvl w:val="0" />
                  <w:numId w:val="21" />
                </w:numPr>
              </w:pPr>
              <w:r><w:t>Top level two</w:t></w:r>
            </w:p>
            ${createCommentedParagraph({
              commentId: '0',
              text: 'This second-level item should continue as b, not restart at a.',
              numbering: { numId: '21', ilvl: '1' },
            })}
          </w:body>
        </w:document>
      `,
      'word/comments.xml': createCommentsXml(['0']),
      'word/numbering.xml': `
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="31">
            <w:lvl w:ilvl="0">
              <w:numFmt w:val="decimal" />
              <w:lvlText w:val="%1." />
            </w:lvl>
            <w:lvl w:ilvl="1">
              <w:numFmt w:val="lowerLetter" />
              <w:lvlText w:val="%1.%2." />
              <w:lvlRestart w:val="0" />
            </w:lvl>
          </w:abstractNum>
          <w:num w:numId="21">
            <w:abstractNumId w:val="31" />
          </w:num>
        </w:numbering>
      `,
    })

    const extracted = await extractParagraphs(file, { ...baseCriteria, comments: true })

    expect(extracted).toHaveLength(1)
    expect(extracted[0]?.numbering).toBe('2.b.')
  })

  it('should allow configurable reset heuristics for carried-forward numbering', async () => {
    const file = await createDocxFile({
      'word/document.xml': `
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:pPr>
                <w:numPr>
                  <w:ilvl w:val="0" />
                  <w:numId w:val="7" />
                </w:numPr>
              </w:pPr>
              <w:r><w:t>1. Definitions</w:t></w:r>
            </w:p>
            ${createCommentedParagraph({
              commentId: '0',
              text: 'Appendix heading should not inherit the clause reference.',
              styleId: 'AppendixHeading',
            })}
          </w:body>
        </w:document>
      `,
      'word/comments.xml': createCommentsXml(['0']),
      'word/numbering.xml': `
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="10">
            <w:lvl w:ilvl="0">
              <w:numFmt w:val="decimal" />
              <w:lvlText w:val="%1." />
            </w:lvl>
          </w:abstractNum>
          <w:num w:numId="7">
            <w:abstractNumId w:val="10" />
          </w:num>
        </w:numbering>
      `,
    })

    const extracted = await extractParagraphs(
      file,
      { ...baseCriteria, comments: true },
      {
        carryForward: {
          shouldReset: ({ styleId }) => styleId === 'AppendixHeading',
        },
      }
    )

    expect(extracted).toHaveLength(1)
    expect(extracted[0]?.numbering).toBeUndefined()
    expect(extracted[0]?.style).toBe('AppendixHeading')
    expect(extracted[0]?.section).toBe(1)
    expect(extracted[0]?.page).toBe(1)
  })
})
