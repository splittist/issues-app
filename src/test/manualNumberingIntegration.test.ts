import { describe, it, expect } from 'vitest';
import { extractParagraphText, detectManualNumbering, validateManualNumbering } from '../numberingUtils';

/**
 * Integration tests for manual numbering detection functionality.
 * These tests demonstrate how the manual numbering detection works
 * in the context of the full application.
 */
describe('Manual Numbering Integration', () => {
  describe('Real-world scenarios', () => {
    it('should detect manual numbering in typical document structures', () => {
      // Simulate paragraph text that would be extracted from real Word documents
      const testParagraphs = [
        {
          text: '1.\tFirst main point with detailed explanation',
          expectedNumbering: '1.',
          description: 'Decimal numbering with tab'
        },
        {
          text: '1.1.\t\tSub-point under first main point',
          expectedNumbering: '1.1.',
          description: 'Multi-level decimal numbering'
        },
        {
          text: 'a.\t   Secondary point in alphabetical sequence',
          expectedNumbering: 'a.',
          description: 'Alphabetical numbering with mixed whitespace'
        },
        {
          text: 'i.\t      Lower level roman numeral point',
          expectedNumbering: 'i.',
          description: 'Roman numeral numbering'
        },
        {
          text: '(1)\tParenthesized numbering style commonly used',
          expectedNumbering: '(1)',
          description: 'Parenthesized decimal numbering'
        },
        {
          text: '(a)\t Alternative parenthesized alphabetical style',
          expectedNumbering: '(a)',
          description: 'Parenthesized alphabetical numbering'
        },
        {
          text: 'This is just regular paragraph text without numbering',
          expectedNumbering: undefined,
          description: 'Regular paragraph without numbering'
        },
        {
          text: '1.Introduction to the topic (this is not numbering)',
          expectedNumbering: undefined,
          description: 'False positive - no proper spacing'
        },
        {
          text: 'www.example.com\tThis looks like numbering but is not',
          expectedNumbering: undefined,
          description: 'False positive - URL pattern'
        }
      ];

      testParagraphs.forEach(testCase => {
        const detected = detectManualNumbering(testCase.text);
        const isValid = detected ? validateManualNumbering(detected, testCase.text) : false;
        const result = isValid ? detected : undefined;

        expect(result).toBe(testCase.expectedNumbering);
        console.log(`✓ ${testCase.description}: "${testCase.text}" → ${result || 'no numbering'}`);
      });
    });

    it('should work correctly with XML paragraph structure', () => {
      // Mock XML paragraph element with manual numbering
      const mockXmlContent = `
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r>
            <w:t>1.</w:t>
            <w:tab/>
            <w:t>This is a manually numbered paragraph in a Word document</w:t>
          </w:r>
        </w:p>
      `;

      const mockDoc = new DOMParser().parseFromString(mockXmlContent, 'text/xml');
      const paragraphElement = mockDoc.documentElement;
      
      // Extract text using the same function used in wordUtils
      const extractedText = extractParagraphText(paragraphElement);
      expect(extractedText).toBe('1.\tThis is a manually numbered paragraph in a Word document');

      // Detect and validate manual numbering
      const detected = detectManualNumbering(extractedText);
      const isValid = detected ? validateManualNumbering(detected, extractedText) : false;
      
      expect(detected).toBe('1.');
      expect(isValid).toBe(true);
    });

    it('should handle mixed automatic and manual numbering scenarios', () => {
      // In real documents, some paragraphs might have automatic numbering
      // while others have manual numbering. The system should handle both.
      
      const scenarios = [
        {
          description: 'Automatic numbering takes precedence',
          automaticNumbering: '2.1.3',
          manualNumberingText: 'a.\tThis has both patterns',
          expectedResult: '2.1.3' // Automatic wins
        },
        {
          description: 'Manual numbering used when automatic is not available',
          automaticNumbering: undefined,
          manualNumberingText: '1.\tThis only has manual numbering',
          expectedResult: '1.' // Manual is used
        },
        {
          description: 'No numbering when neither is valid',
          automaticNumbering: undefined,
          manualNumberingText: 'Just regular text',
          expectedResult: undefined // No numbering
        }
      ];

      scenarios.forEach(scenario => {
        // Simulate the logic from wordUtils.ts
        let finalNumbering = scenario.automaticNumbering;
        
        if (!finalNumbering) {
          const detected = detectManualNumbering(scenario.manualNumberingText);
          if (detected && validateManualNumbering(detected, scenario.manualNumberingText)) {
            finalNumbering = detected;
          }
        }

        expect(finalNumbering).toBe(scenario.expectedResult);
        console.log(`✓ ${scenario.description}: ${finalNumbering || 'no numbering'}`);
      });
    });

    it('should demonstrate the improvement over previous behavior', () => {
      // Before: Documents with manual numbering would show "Sect X, p Y" references
      // After: Documents with manual numbering will show the actual numbering
      
      const manuallyNumberedParagraphs = [
        '1.\tFirst major section',
        '1.1.\tSubsection of first major section', 
        '1.2.\tAnother subsection',
        '2.\tSecond major section',
        'a.\tAlphabetical subsection',
        'b.\tAnother alphabetical subsection',
        'i.\tRoman numeral detail',
        'ii.\tContinued roman detail'
      ];

      const detectedNumberings = manuallyNumberedParagraphs.map(text => {
        const detected = detectManualNumbering(text);
        return detected && validateManualNumbering(detected, text) ? detected : undefined;
      });

      // Verify that we correctly detected all the manual numbering
      expect(detectedNumberings).toEqual([
        '1.',
        '1.1.',
        '1.2.',
        '2.',
        'a.',
        'b.',
        'i.',
        'ii.'
      ]);

      console.log('Manual numbering detection results:');
      manuallyNumberedParagraphs.forEach((text, index) => {
        const numbering = detectedNumberings[index];
        const beforeReference = `Sect 1, p ${index + 1}`; // What it would have been before
        const afterReference = numbering || beforeReference; // What it is now
        
        console.log(`  "${text}" → Before: "${beforeReference}", After: "${afterReference}"`);
      });
    });
  });

  describe('Edge cases and robustness', () => {
    it('should handle various whitespace patterns correctly', () => {
      const whitespaceVariations = [
        { text: '1.\tSingle tab', expected: '1.' },
        { text: '1.\t\tDouble tab', expected: '1.' },
        { text: '1.   Three spaces', expected: '1.' },
        { text: '1.    Four spaces', expected: '1.' },
        { text: '1. \t Mixed space and tab', expected: '1.' },
        { text: '1.\t  Tab then spaces', expected: '1.' },
        { text: '1.  \t Spaces then tab', expected: '1.' }
      ];

      whitespaceVariations.forEach(variation => {
        const detected = detectManualNumbering(variation.text);
        const isValid = detected ? validateManualNumbering(detected, variation.text) : false;
        const result = isValid ? detected : undefined;
        
        expect(result).toBe(variation.expected);
      });
    });

    it('should reject common false positive patterns', () => {
      const falsePositives = [
        'version 1.2.3.4 of the software',
        'www.example.com website',
        'IP address 192.168.1.1 configuration',
        'email@domain.com contact',
        'file.txt document',
        '1.This has no space after the dot',
        'a.No space here either',
        '1. X', // Too short content
        'a. Y'  // Too short content
      ];

      falsePositives.forEach(text => {
        const detected = detectManualNumbering(text);
        const isValid = detected ? validateManualNumbering(detected, text) : false;
        const result = isValid ? detected : undefined;
        
        expect(result).toBeUndefined();
      });
    });
  });
});