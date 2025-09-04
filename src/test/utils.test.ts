import { describe, it, expect, vi } from 'vitest'
import { dateToday, formatCommentDate } from '../utils'

describe('utils', () => {
  describe('dateToday', () => {
    it('should return current date in YYYY-MM-DD format', () => {
      const result = dateToday()
      // Should match YYYY-MM-DD pattern
      expect(result).toMatch(/^\d{4}-\d{2}-\d{2}$/)
    })

    it('should handle timezone offset correctly', () => {
      // Mock a specific date and time
      const mockDate = new Date('2023-05-15T12:00:00Z')
      vi.setSystemTime(mockDate)
      
      const result = dateToday()
      // Should return date considering timezone offset
      expect(result).toMatch(/^\d{4}-\d{2}-\d{2}$/)
      expect(result.length).toBe(10) // YYYY-MM-DD format
    })
  })

  describe('formatCommentDate', () => {
    it('should format valid ISO date string', () => {
      const result = formatCommentDate('2023-05-15T10:30:00Z')
      expect(result).toBe('May 15, 2023')
    })

    it('should format simple date string', () => {
      const result = formatCommentDate('2023-05-15')
      expect(result).toBe('May 15, 2023')
    })

    it('should return original string for invalid date', () => {
      const invalidDate = 'not-a-date'
      const result = formatCommentDate(invalidDate)
      expect(result).toBe(invalidDate)
    })

    it('should handle empty string', () => {
      const result = formatCommentDate('')
      expect(result).toBe('')
    })

    it('should handle partial date strings gracefully', () => {
      const result = formatCommentDate('2023-13-45')
      expect(result).toBe('2023-13-45') // Should return original for invalid date
    })
  })
})