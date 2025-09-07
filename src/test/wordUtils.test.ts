import { describe, it, expect } from 'vitest'
import { ExtractedParagraph } from '../types'
import { Paragraph } from 'docx'

describe('wordUtils', () => {
  describe('reference generation for headers and footers', () => {
    it('should generate correct reference for header paragraphs', () => {
      // Create a mock ExtractedParagraph for a header
      const headerParagraph: ExtractedParagraph = {
        paragraph: new Paragraph('Test header content'),
        comments: [],
        section: 1,
        page: undefined, // Headers don't have page numbers
        numbering: undefined,
        style: undefined,
        source: 'header'
      }

      // Test that the reference should be "Sect 1, Header" not "Sect 1, p undefined"
      const expectedReference = 'Sect 1, Header'
      const actualReference = headerParagraph.numbering || 
        (headerParagraph.source === 'header' || headerParagraph.source === 'footer' 
          ? `Sect ${headerParagraph.section}, ${headerParagraph.source.charAt(0).toUpperCase() + headerParagraph.source.slice(1)}` 
          : `Sect ${headerParagraph.section}, p ${headerParagraph.page}`)
      
      expect(actualReference).toBe(expectedReference)
    })

    it('should generate correct reference for footer paragraphs', () => {
      // Create a mock ExtractedParagraph for a footer
      const footerParagraph: ExtractedParagraph = {
        paragraph: new Paragraph('Test footer content'),
        comments: [],
        section: 2,
        page: undefined, // Footers don't have page numbers
        numbering: undefined,
        style: undefined,
        source: 'footer'
      }

      // Test that the reference should be "Sect 2, Footer" not "Sect 2, p undefined"
      const expectedReference = 'Sect 2, Footer'
      const actualReference = footerParagraph.numbering || 
        (footerParagraph.source === 'header' || footerParagraph.source === 'footer' 
          ? `Sect ${footerParagraph.section}, ${footerParagraph.source.charAt(0).toUpperCase() + footerParagraph.source.slice(1)}` 
          : `Sect ${footerParagraph.section}, p ${footerParagraph.page}`)
      
      expect(actualReference).toBe(expectedReference)
    })

    it('should generate correct reference for document paragraphs', () => {
      // Create a mock ExtractedParagraph for a document
      const documentParagraph: ExtractedParagraph = {
        paragraph: new Paragraph('Test document content'),
        comments: [],
        section: 1,
        page: 5,
        numbering: undefined,
        style: undefined,
        source: 'document'
      }

      // Test that the reference should remain "Sect 1, p 5" for documents
      const expectedReference = 'Sect 1, p 5'
      const actualReference = documentParagraph.numbering || 
        (documentParagraph.source === 'header' || documentParagraph.source === 'footer' 
          ? `Sect ${documentParagraph.section}, ${documentParagraph.source.charAt(0).toUpperCase() + documentParagraph.source.slice(1)}` 
          : `Sect ${documentParagraph.section}, p ${documentParagraph.page}`)
      
      expect(actualReference).toBe(expectedReference)
    })

    it('should use numbering when available instead of section/page reference', () => {
      // Create a mock ExtractedParagraph with numbering
      const numberedParagraph: ExtractedParagraph = {
        paragraph: new Paragraph('Test numbered content'),
        comments: [],
        section: 1,
        page: undefined,
        numbering: '1.2.3',
        style: undefined,
        source: 'header'
      }

      // Test that numbering takes precedence
      const expectedReference = '1.2.3'
      const actualReference = numberedParagraph.numbering || 
        (numberedParagraph.source === 'header' || numberedParagraph.source === 'footer' 
          ? `Sect ${numberedParagraph.section}, ${numberedParagraph.source.charAt(0).toUpperCase() + numberedParagraph.source.slice(1)}` 
          : `Sect ${numberedParagraph.section}, p ${numberedParagraph.page}`)
      
      expect(actualReference).toBe(expectedReference)
    })
  })
})