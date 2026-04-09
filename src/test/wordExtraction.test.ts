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
})
