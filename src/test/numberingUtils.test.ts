import { describe, it, expect } from 'vitest'
import {
  toLowerLetter,
  toUpperLetter,
  toLowerRoman,
  toUpperRoman,
  toOrdinal,
  toCardinalText,
  toOrdinalText,
  toNumberInDash,
  formatNumber,
  processLvlText,
  initializeCounters,
  updateCounters,
  buildNumberingMaps,
  buildStyleMaps,
  resolveStyleNumbering,
  extractParagraphStyle
} from '../numberingUtils'

describe('numberingUtils', () => {
  describe('toLowerLetter', () => {
    it('should convert numbers to lowercase letters', () => {
      expect(toLowerLetter(1)).toBe('a')
      expect(toLowerLetter(2)).toBe('b')
      expect(toLowerLetter(26)).toBe('z')
      expect(toLowerLetter(27)).toBe('aa')
      expect(toLowerLetter(28)).toBe('ab')
      expect(toLowerLetter(52)).toBe('az')
      expect(toLowerLetter(53)).toBe('ba')
    })

    it('should handle edge cases', () => {
      expect(toLowerLetter(0)).toBe('')
      expect(toLowerLetter(-1)).toBe('')
      expect(toLowerLetter(702)).toBe('zz') // 26*26 + 26
      expect(toLowerLetter(703)).toBe('aaa')
    })

    it('should handle very large numbers', () => {
      expect(toLowerLetter(18278)).toBe('zzz') // 26^3 + 26^2 + 26
      expect(toLowerLetter(18279)).toBe('aaaa')
    })
  })

  describe('toUpperLetter', () => {
    it('should convert numbers to uppercase letters', () => {
      expect(toUpperLetter(1)).toBe('A')
      expect(toUpperLetter(2)).toBe('B')
      expect(toUpperLetter(26)).toBe('Z')
      expect(toUpperLetter(27)).toBe('AA')
      expect(toUpperLetter(28)).toBe('AB')
      expect(toUpperLetter(52)).toBe('AZ')
      expect(toUpperLetter(53)).toBe('BA')
    })

    it('should handle edge cases', () => {
      expect(toUpperLetter(0)).toBe('')
      expect(toUpperLetter(-1)).toBe('')
      expect(toUpperLetter(702)).toBe('ZZ')
      expect(toUpperLetter(703)).toBe('AAA')
    })
  })

  describe('toLowerRoman', () => {
    it('should convert numbers to lowercase Roman numerals', () => {
      expect(toLowerRoman(1)).toBe('i')
      expect(toLowerRoman(2)).toBe('ii')
      expect(toLowerRoman(3)).toBe('iii')
      expect(toLowerRoman(4)).toBe('iv')
      expect(toLowerRoman(5)).toBe('v')
      expect(toLowerRoman(9)).toBe('ix')
      expect(toLowerRoman(10)).toBe('x')
      expect(toLowerRoman(40)).toBe('xl')
      expect(toLowerRoman(50)).toBe('l')
      expect(toLowerRoman(90)).toBe('xc')
      expect(toLowerRoman(100)).toBe('c')
      expect(toLowerRoman(400)).toBe('cd')
      expect(toLowerRoman(500)).toBe('d')
      expect(toLowerRoman(900)).toBe('cm')
      expect(toLowerRoman(1000)).toBe('m')
    })

    it('should handle complex Roman numerals', () => {
      expect(toLowerRoman(44)).toBe('xliv')
      expect(toLowerRoman(1994)).toBe('mcmxciv')
      expect(toLowerRoman(2023)).toBe('mmxxiii')
    })

    it('should handle edge cases', () => {
      expect(toLowerRoman(0)).toBe('')
      expect(toLowerRoman(-1)).toBe('')
    })

    it('should handle very large numbers', () => {
      expect(toLowerRoman(3999)).toBe('mmmcmxcix') // Largest standard Roman numeral
      expect(toLowerRoman(4000)).toBe('mmmm') // Beyond standard range
    })
  })

  describe('toUpperRoman', () => {
    it('should convert numbers to uppercase Roman numerals', () => {
      expect(toUpperRoman(1)).toBe('I')
      expect(toUpperRoman(2)).toBe('II')
      expect(toUpperRoman(3)).toBe('III')
      expect(toUpperRoman(4)).toBe('IV')
      expect(toUpperRoman(5)).toBe('V')
      expect(toUpperRoman(9)).toBe('IX')
      expect(toUpperRoman(10)).toBe('X')
      expect(toUpperRoman(44)).toBe('XLIV')
      expect(toUpperRoman(1994)).toBe('MCMXCIV')
    })

    it('should handle edge cases', () => {
      expect(toUpperRoman(0)).toBe('')
      expect(toUpperRoman(-1)).toBe('')
    })
  })

  describe('toOrdinal', () => {
    it('should convert numbers to ordinals', () => {
      expect(toOrdinal(1)).toBe('1st')
      expect(toOrdinal(2)).toBe('2nd')
      expect(toOrdinal(3)).toBe('3rd')
      expect(toOrdinal(4)).toBe('4th')
      expect(toOrdinal(11)).toBe('11th')
      expect(toOrdinal(12)).toBe('12th')
      expect(toOrdinal(13)).toBe('13th')
      expect(toOrdinal(21)).toBe('21st')
      expect(toOrdinal(22)).toBe('22nd')
      expect(toOrdinal(23)).toBe('23rd')
      expect(toOrdinal(101)).toBe('101st')
      expect(toOrdinal(102)).toBe('102nd')
      expect(toOrdinal(103)).toBe('103rd')
      expect(toOrdinal(111)).toBe('111th')
      expect(toOrdinal(112)).toBe('112th')
      expect(toOrdinal(113)).toBe('113th')
    })

    it('should handle edge cases', () => {
      expect(toOrdinal(0)).toBe('0th')
      expect(toOrdinal(-1)).toBe('-1th')
    })

    it('should handle complex ordinal patterns', () => {
      expect(toOrdinal(121)).toBe('121st')
      expect(toOrdinal(122)).toBe('122nd')
      expect(toOrdinal(123)).toBe('123rd')
      expect(toOrdinal(124)).toBe('124th')
    })
  })

  describe('toCardinalText', () => {
    it('should convert single digits to text', () => {
      expect(toCardinalText(0)).toBe('zero')
      expect(toCardinalText(1)).toBe('one')
      expect(toCardinalText(2)).toBe('two')
      expect(toCardinalText(9)).toBe('nine')
    })

    it('should convert teens to text', () => {
      expect(toCardinalText(10)).toBe('ten')
      expect(toCardinalText(11)).toBe('eleven')
      expect(toCardinalText(12)).toBe('twelve')
      expect(toCardinalText(13)).toBe('thirteen')
      expect(toCardinalText(19)).toBe('nineteen')
    })

    it('should convert tens to text', () => {
      expect(toCardinalText(20)).toBe('twenty')
      expect(toCardinalText(21)).toBe('twenty-one')
      expect(toCardinalText(30)).toBe('thirty')
      expect(toCardinalText(45)).toBe('forty-five')
      expect(toCardinalText(99)).toBe('ninety-nine')
    })

    it('should convert hundreds to text', () => {
      expect(toCardinalText(100)).toBe('one hundred')
      expect(toCardinalText(101)).toBe('one hundred one')
      expect(toCardinalText(115)).toBe('one hundred fifteen')
      expect(toCardinalText(123)).toBe('one hundred twenty-three')
      expect(toCardinalText(999)).toBe('nine hundred ninety-nine')
    })

    it('should handle negative numbers', () => {
      expect(toCardinalText(-1)).toBe('negative one')
      expect(toCardinalText(-42)).toBe('negative forty-two')
    })

    it('should fallback for large numbers', () => {
      expect(toCardinalText(1000)).toBe('1000')
      expect(toCardinalText(1500)).toBe('1500')
    })
  })

  describe('toOrdinalText', () => {
    it('should convert numbers 1-20 to ordinal text', () => {
      expect(toOrdinalText(1)).toBe('first')
      expect(toOrdinalText(2)).toBe('second')
      expect(toOrdinalText(3)).toBe('third')
      expect(toOrdinalText(4)).toBe('fourth')
      expect(toOrdinalText(5)).toBe('fifth')
      expect(toOrdinalText(8)).toBe('eighth')
      expect(toOrdinalText(9)).toBe('ninth')
      expect(toOrdinalText(12)).toBe('twelfth')
      expect(toOrdinalText(20)).toBe('twentieth')
    })

    it('should convert larger numbers to ordinal text', () => {
      expect(toOrdinalText(21)).toBe('twenty-first')
      expect(toOrdinalText(22)).toBe('twenty-second')
      expect(toOrdinalText(23)).toBe('twenty-third')
      expect(toOrdinalText(25)).toBe('twenty-fifth')
      expect(toOrdinalText(28)).toBe('twenty-eighth')
      expect(toOrdinalText(29)).toBe('twenty-ninth')
      expect(toOrdinalText(30)).toBe('thirtieth')
      expect(toOrdinalText(32)).toBe('thirty-second')
    })

    it('should handle hundreds', () => {
      expect(toOrdinalText(101)).toBe('one hundred first')
      expect(toOrdinalText(102)).toBe('one hundred second')
      expect(toOrdinalText(103)).toBe('one hundred third')
      expect(toOrdinalText(105)).toBe('one hundred fifth')
      expect(toOrdinalText(112)).toBe('one hundred twelfth')
    })

    it('should handle edge cases for ordinal text', () => {
      expect(toOrdinalText(0)).toBe('zeroth')
      expect(toOrdinalText(-1)).toBe('negative first')
    })
  })

  describe('toNumberInDash', () => {
    it('should format numbers with dashes', () => {
      expect(toNumberInDash(1)).toBe('- 1 -')
      expect(toNumberInDash(42)).toBe('- 42 -')
      expect(toNumberInDash(0)).toBe('- 0 -')
      expect(toNumberInDash(-5)).toBe('- -5 -')
    })
  })

  describe('formatNumber', () => {
    it('should dispatch to correct format function', () => {
      expect(formatNumber(5, 'decimal')).toBe('5')
      expect(formatNumber(5, 'lowerLetter')).toBe('e')
      expect(formatNumber(5, 'upperLetter')).toBe('E')
      expect(formatNumber(5, 'lowerRoman')).toBe('v')
      expect(formatNumber(5, 'upperRoman')).toBe('V')
      expect(formatNumber(5, 'bullet')).toBe('•')
      expect(formatNumber(5, 'ordinal')).toBe('5th')
      expect(formatNumber(5, 'cardinalText')).toBe('five')
      expect(formatNumber(5, 'ordinalText')).toBe('fifth')
      expect(formatNumber(5, 'numberInDash')).toBe('- 5 -')
    })

    it('should default to decimal for unknown formats', () => {
      expect(formatNumber(5, 'unknownFormat')).toBe('5')
      expect(formatNumber(5, '')).toBe('5')
    })
  })

  describe('initializeCounters', () => {
    it('should create array of zeros', () => {
      expect(initializeCounters(3)).toEqual([0, 0, 0])
      expect(initializeCounters(5)).toEqual([0, 0, 0, 0, 0])
      expect(initializeCounters(0)).toEqual([])
    })
  })

  describe('updateCounters', () => {
    it('should increment specified level and reset lower levels', () => {
      const counters = [1, 2, 3, 4];
      const updated = updateCounters(counters, 1);
      expect(updated).toEqual([1, 3, 0, 0]);
    })

    it('should handle level 0', () => {
      const counters = [1, 2, 3];
      const updated = updateCounters(counters, 0);
      expect(updated).toEqual([2, 0, 0]);
    })

    it('should handle highest level', () => {
      const counters = [1, 2, 3];
      const updated = updateCounters(counters, 2);
      expect(updated).toEqual([1, 2, 4]);
    })

    it('should return the same array reference', () => {
      const counters = [1, 2, 3];
      const updated = updateCounters(counters, 1);
      expect(updated).toBe(counters); // Same reference
    })
  })

  describe('buildNumberingMaps', () => {
    it('should handle empty numbering document', () => {
      const mockDoc = {
        getElementsByTagName: () => []
      } as unknown as globalThis.Document;

      const result = buildNumberingMaps(mockDoc);
      expect(result.numIdToAbstractNumId.size).toBe(0);
      expect(result.abstractNumIdToFormat.size).toBe(0);
    })

    it('should parse simple numbering document', () => {
      const mockNumElement = {
        getAttribute: (attr: string) => attr === 'w:numId' ? '1' : null,
        getElementsByTagName: () => [
          { getAttribute: (attr: string) => attr === 'w:val' ? '0' : null }
        ]
      };

      const mockAbstractNumElement = {
        getAttribute: (attr: string) => attr === 'w:abstractNumId' ? '0' : null,
        getElementsByTagName: () => [
          {
            getElementsByTagName: () => [
              { getAttribute: (attr: string) => attr === 'w:val' ? 'decimal' : null }
            ]
          }
        ]
      };

      const mockDoc = {
        getElementsByTagName: (tagName: string) => {
          if (tagName === 'w:num') return [mockNumElement];
          if (tagName === 'w:abstractNum') return [mockAbstractNumElement];
          return [];
        }
      } as unknown as globalThis.Document;

      const result = buildNumberingMaps(mockDoc);
      expect(result.numIdToAbstractNumId.get('1')).toBe('0');
    })
  })

  describe('buildStyleMaps', () => {
    it('should handle empty styles document', () => {
      const mockDoc = {
        getElementsByTagName: () => []
      } as unknown as globalThis.Document;

      const result = buildStyleMaps(mockDoc);
      expect(result.size).toBe(0);
    })

    it('should parse simple style without numbering', () => {
      const mockStyleElement = {
        getAttribute: (attr: string) => attr === 'w:styleId' ? 'heading1' : null,
        getElementsByTagName: (tagName: string) => {
          if (tagName === 'w:name') return [
            { getAttribute: (attr: string) => attr === 'w:val' ? 'Heading 1' : null }
          ];
          return [];
        }
      };

      const mockDoc = {
        getElementsByTagName: (tagName: string) => 
          tagName === 'w:style' ? [mockStyleElement] : []
      } as unknown as globalThis.Document;

      const result = buildStyleMaps(mockDoc);
      expect(result.size).toBe(1);
      expect(result.get('heading1')?.name).toBe('Heading 1');
      expect(result.get('heading1')?.numbering).toBeUndefined();
    })
  })

  describe('resolveStyleNumbering', () => {
    it('should return undefined for non-existent style', () => {
      const styles = new Map();
      const result = resolveStyleNumbering('nonexistent', styles);
      expect(result).toBeUndefined();
    })

    it('should return direct numbering', () => {
      const styles = new Map([
        ['heading1', { 
          id: 'heading1', 
          name: 'Heading 1', 
          numbering: { numId: '1', ilvl: '0' } 
        }]
      ]);
      
      const result = resolveStyleNumbering('heading1', styles);
      expect(result).toEqual({ numId: '1', ilvl: '0' });
    })

    it('should resolve through inheritance chain', () => {
      const styles = new Map([
        ['child', { 
          id: 'child', 
          name: 'Child', 
          basedOn: 'parent' 
        }],
        ['parent', { 
          id: 'parent', 
          name: 'Parent', 
          numbering: { numId: '2', ilvl: '1' } 
        }]
      ]);
      
      const result = resolveStyleNumbering('child', styles);
      expect(result).toEqual({ numId: '2', ilvl: '1' });
    })

    it('should handle circular references', () => {
      const styles = new Map([
        ['a', { id: 'a', name: 'A', basedOn: 'b' }],
        ['b', { id: 'b', name: 'B', basedOn: 'a' }]
      ]);
      
      const result = resolveStyleNumbering('a', styles);
      expect(result).toBeUndefined();
    })
  })

  describe('extractParagraphStyle', () => {
    it('should return undefined for paragraph without style', () => {
      const mockElement = {
        getElementsByTagName: () => []
      } as unknown as Element;

      const result = extractParagraphStyle(mockElement);
      expect(result).toBeUndefined();
    })

    it('should extract style ID', () => {
      const mockPStyleElement = {
        getAttribute: (attr: string) => attr === 'w:val' ? 'heading1' : null
      };

      const mockPPrElement = {
        getElementsByTagName: (tagName: string) => 
          tagName === 'w:pStyle' ? [mockPStyleElement] : []
      };

      const mockElement = {
        getElementsByTagName: (tagName: string) => 
          tagName === 'w:pPr' ? [mockPPrElement] : []
      } as unknown as Element;

      const result = extractParagraphStyle(mockElement);
      expect(result).toBe('heading1');
    })
  })

  describe('processLvlText', () => {
    it('should process simple lvlText template with single level', () => {
      const template = '%1.';
      const numbers = ['1'];
      const result = processLvlText(template, numbers);
      expect(result).toBe('1.');
    })

    it('should process multi-level lvlText template', () => {
      const template = '%1.%2(%3)';
      const numbers = ['2', '1', 'i'];
      const result = processLvlText(template, numbers);
      expect(result).toBe('2.1(i)');
    })

    it('should handle lvlText template with different separators', () => {
      const template = '%1)';
      const numbers = ['3'];
      const result = processLvlText(template, numbers);
      expect(result).toBe('3)');
    })

    it('should handle empty template by joining with dots', () => {
      const template = '';
      const numbers = ['1', '2', '3'];
      const result = processLvlText(template, numbers);
      expect(result).toBe('1.2.3');
    })

    it('should handle template with extra placeholders', () => {
      const template = '%1.%2.%3.%4';
      const numbers = ['1', '2'];
      const result = processLvlText(template, numbers);
      expect(result).toBe('1.2.%3.%4');
    })
  })

  describe('full numbering integration', () => {
    it('should demonstrate fully qualified numbering with lvlText templates', () => {
      // Mock data structure simulating Word document numbering
      const formats = [
        { numFmt: 'decimal', lvlText: '%1.' },           // Level 0: "1."
        { numFmt: 'decimal', lvlText: '%1.%2.' },        // Level 1: "1.1."
        { numFmt: 'lowerRoman', lvlText: '%1.%2(%3)' },  // Level 2: "1.1(i)"
      ];
      
      // Simulate counters at level 2 with values [2, 1, 3]
      const counters = [2, 1, 3];
      const currentLevel = 2;
      
      // Format numbers according to their level's numFmt
      const formattedNumbers = counters.slice(0, currentLevel + 1)
        .map((num, index) => {
          const fmt = formats[index]?.numFmt || 'decimal';
          return formatNumber(num, fmt);
        });
      
      // Should be ['2', '1', 'iii']
      expect(formattedNumbers).toEqual(['2', '1', 'iii']);
      
      // Use current level's lvlText template
      const result = processLvlText(formats[currentLevel].lvlText, formattedNumbers);
      
      // Should produce fully qualified numbering: "2.1(iii)"
      expect(result).toBe('2.1(iii)');
    })

    it('should handle different numbering formats properly', () => {
      const formats = [
        { numFmt: 'upperLetter', lvlText: '%1)' },        // Level 0: "A)"
        { numFmt: 'lowerLetter', lvlText: '%1)%2)' },     // Level 1: "A)a)"
        { numFmt: 'decimal', lvlText: '%1)%2)%3.' },      // Level 2: "A)a)1."
      ];
      
      const counters = [3, 2, 5]; // C, b, 5
      const currentLevel = 2;
      
      const formattedNumbers = counters.slice(0, currentLevel + 1)
        .map((num, index) => {
          const fmt = formats[index]?.numFmt || 'decimal';
          return formatNumber(num, fmt);
        });
      
      expect(formattedNumbers).toEqual(['C', 'b', '5']);
      
      const result = processLvlText(formats[currentLevel].lvlText, formattedNumbers);
      expect(result).toBe('C)b)5.');
    })
  })
})