import { describe, expect, it } from 'vitest'
import {
  defaultCarryForwardPolicy,
  resolveCarryForwardPolicy,
} from '../wordNumberingPolicy'

describe('wordNumberingPolicy', () => {
  it('should enable carry-forward by default', () => {
    expect(defaultCarryForwardPolicy.enabled).toBe(true)
    expect(defaultCarryForwardPolicy.shouldReset({
      currentPage: 1,
      currentSection: 1,
      paragraphElement: new DOMParser().parseFromString('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>', 'text/xml').documentElement,
      paragraphText: 'Example paragraph',
      previousNumberingInfo: '1.',
      styleId: undefined,
    })).toBe(false)
  })

  it('should preserve defaults when no options are provided', () => {
    const policy = resolveCarryForwardPolicy()

    expect(policy.enabled).toBe(true)
    expect(policy.shouldReset({
      currentPage: 2,
      currentSection: 3,
      paragraphElement: new DOMParser().parseFromString('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>', 'text/xml').documentElement,
      paragraphText: 'Another paragraph',
      previousNumberingInfo: '2.1',
      styleId: 'BodyText',
    })).toBe(false)
  })

  it('should allow callers to override enabled state and reset logic', () => {
    const policy = resolveCarryForwardPolicy({
      carryForward: {
        enabled: false,
        shouldReset: ({ styleId, paragraphText }) => styleId === 'AppendixHeading' || paragraphText.includes('SCHEDULE'),
      },
    })

    expect(policy.enabled).toBe(false)
    expect(policy.shouldReset({
      currentPage: 1,
      currentSection: 1,
      paragraphElement: new DOMParser().parseFromString('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>', 'text/xml').documentElement,
      paragraphText: 'SCHEDULE 1',
      previousNumberingInfo: '4.',
      styleId: 'BodyText',
    })).toBe(true)
  })
})
