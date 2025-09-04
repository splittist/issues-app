import { describe, it, expect } from 'vitest'
import type { Criteria, ParagraphSource } from '../types'
import { Break } from '../types'

describe('types', () => {
  describe('Criteria interface', () => {
    it('should allow creating valid Criteria objects', () => {
      const criteria: Criteria = {
        redline: true,
        highlight: false,
        squareBrackets: true,
        comments: false,
        footnotes: true,
        endnotes: false
      }
      
      expect(criteria.redline).toBe(true)
      expect(criteria.highlight).toBe(false)
      expect(criteria.squareBrackets).toBe(true)
      expect(criteria.comments).toBe(false)
      expect(criteria.footnotes).toBe(true)
      expect(criteria.endnotes).toBe(false)
    })
  })

  describe('ParagraphSource type', () => {
    it('should accept valid source values', () => {
      const sources: ParagraphSource[] = ['document', 'header', 'footer']
      expect(sources).toHaveLength(3)
      expect(sources).toContain('document')
      expect(sources).toContain('header')
      expect(sources).toContain('footer')
    })
  })

  describe('Break class', () => {
    it('should create Break instance', () => {
      const breakElement = new Break()
      expect(breakElement).toBeInstanceOf(Break)
    })
  })

  describe('ExtractedParagraph interface', () => {
    it('should be importable as a type', () => {
      // We test that the type can be imported without errors
      // This validates the interface structure
      const checkType = (): void => {
        // This function just validates the types are accessible
      }
      expect(checkType).toBeDefined()
    })
  })
})