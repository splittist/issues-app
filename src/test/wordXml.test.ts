import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { Document, Packer, Paragraph } from 'docx'
import { buildDocumentParagraph, buildTextRun } from '../wordXml'

const paragraphToXml = async (paragraph: Paragraph): Promise<string> => {
  const doc = new Document({
    sections: [{ children: [paragraph] }]
  })
  const buffer = await Packer.toBuffer(doc)
  const zip = await JSZip.loadAsync(buffer)
  return await zip.file('word/document.xml')!.async('string')
}

const getXmlElement = (xml: string, selector: string): Element => {
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  const element = doc.querySelector(selector)
  if (!element) {
    throw new Error(`Missing selector: ${selector}`)
  }

  return element
}

describe('wordXml', () => {
  describe('buildTextRun', () => {
    it('should preserve text, tabs, and annotation references in a single run', async () => {
      const runElement = getXmlElement(`
        <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:t>Hello</w:t>
          <w:tab />
          <w:t>world</w:t>
          <w:commentReference w:id="7" />
          <w:t>again</w:t>
        </w:r>
      `, 'w\\:r')

      const paragraph = new Paragraph({ children: buildTextRun(runElement) })
      const xml = await paragraphToXml(paragraph)

      expect(xml).toContain('Hello')
      expect(xml).toContain('world')
      expect(xml).toContain('[Cmt 7]')
      expect(xml).toContain('again')
      expect(xml).toContain('CommentAnchor')
    })
  })

  describe('buildDocumentParagraph', () => {
    it('should keep insertion and deletion text with the expected styles', async () => {
      const paragraphElement = getXmlElement(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r>
            <w:t>Start </w:t>
          </w:r>
          <w:del>
            <w:r>
              <w:delText>Removed</w:delText>
            </w:r>
          </w:del>
          <w:ins>
            <w:r>
              <w:t>Added</w:t>
            </w:r>
          </w:ins>
        </w:p>
      `, 'w\\:p')

      const paragraph = buildDocumentParagraph(paragraphElement)
      const xml = await paragraphToXml(paragraph)
      const serializedParagraph = JSON.stringify(paragraph)

      expect(xml).toContain('Start ')
      expect(xml).toContain('Removed')
      expect(xml).toContain('Added')
      expect(serializedParagraph).toContain('Deletion')
      expect(serializedParagraph).toContain('Insertion')
    })

    it('should keep move markers with their dedicated styles', async () => {
      const paragraphElement = getXmlElement(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:moveFrom>
            <w:r>
              <w:t>Old location</w:t>
            </w:r>
          </w:moveFrom>
          <w:moveTo>
            <w:r>
              <w:t>New location</w:t>
            </w:r>
          </w:moveTo>
        </w:p>
      `, 'w\\:p')

      const paragraph = buildDocumentParagraph(paragraphElement)
      const xml = await paragraphToXml(paragraph)
      const serializedParagraph = JSON.stringify(paragraph)

      expect(xml).toContain('Old location')
      expect(xml).toContain('New location')
      expect(serializedParagraph).toContain('MoveFrom')
      expect(serializedParagraph).toContain('MoveTo')
    })

    it('should keep revisions nested inside transparent wrapper elements', async () => {
      const paragraphElement = getXmlElement(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r>
            <w:t>Start </w:t>
          </w:r>
          <w:hyperlink w:anchor="_Ref1">
            <w:bookmarkStart w:id="1" w:name="_Ref1" />
            <w:del>
              <w:r>
                <w:delText>Nested removed</w:delText>
              </w:r>
            </w:del>
            <w:bookmarkEnd w:id="1" />
          </w:hyperlink>
          <w:customXml>
            <w:ins>
              <w:r>
                <w:t>Nested added</w:t>
              </w:r>
            </w:ins>
          </w:customXml>
        </w:p>
      `, 'w\\:p')

      const paragraph = buildDocumentParagraph(paragraphElement)
      const xml = await paragraphToXml(paragraph)
      const serializedParagraph = JSON.stringify(paragraph)

      expect(xml).toContain('Start ')
      expect(xml).toContain('Nested removed')
      expect(xml).toContain('Nested added')
      expect(serializedParagraph).toContain('Deletion')
      expect(serializedParagraph).toContain('Insertion')
    })
  })
})
