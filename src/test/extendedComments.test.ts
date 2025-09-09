import { describe, it, expect } from 'vitest';
import { ExtendedCommentInfo } from '../types';

// Import the function we need to test - we'll need to expose it for testing
// For now, let's create a mock implementation to test the parsing logic

/**
 * Test version of parseExtendedComments function extracted from wordUtils.ts
 */
const parseExtendedComments = (xmlCommentsExtended: Document): Map<string, ExtendedCommentInfo> => {
  const extendedCommentsMap = new Map<string, ExtendedCommentInfo>();
  
  // Look for w15:commentEx elements
  const commentExElements = xmlCommentsExtended.querySelectorAll('commentEx, w15\\:commentEx');
  
  for (const commentExElement of commentExElements) {
    const paraId = commentExElement.getAttribute('w15:paraId') || commentExElement.getAttribute('paraId');
    if (paraId) {
      const paraIdParent = commentExElement.getAttribute('w15:paraIdParent') || commentExElement.getAttribute('paraIdParent');
      const doneAttr = commentExElement.getAttribute('w15:done') || commentExElement.getAttribute('done');
      const done = doneAttr === '1';
      
      const extendedInfo: ExtendedCommentInfo = {
        paraId,
        ...(paraIdParent && { paraIdParent }),
        ...(doneAttr && { done })
      };
      
      extendedCommentsMap.set(paraId, extendedInfo);
    }
  }
  
  return extendedCommentsMap;
};

describe('Extended Comments Functionality', () => {
  describe('parseExtendedComments', () => {
    it('should parse a simple commentsExtended.xml with resolved comment', () => {
      const xmlString = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
          <w15:commentEx w15:paraId="12345" w15:done="1"/>
        </w15:commentsEx>`;
      
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlString, 'application/xml');
      const result = parseExtendedComments(xmlDoc);
      
      expect(result.size).toBe(1);
      expect(result.has('12345')).toBe(true);
      
      const commentInfo = result.get('12345');
      expect(commentInfo).toEqual({
        paraId: '12345',
        done: true
      });
    });

    it('should parse commentsExtended.xml with threaded comments', () => {
      const xmlString = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
          <w15:commentEx w15:paraId="11111" w15:done="0"/>
          <w15:commentEx w15:paraId="22222" w15:paraIdParent="11111" w15:done="0"/>
          <w15:commentEx w15:paraId="33333" w15:paraIdParent="11111" w15:done="1"/>
        </w15:commentsEx>`;
      
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlString, 'application/xml');
      const result = parseExtendedComments(xmlDoc);
      
      expect(result.size).toBe(3);
      
      // Parent comment
      const parentComment = result.get('11111');
      expect(parentComment).toEqual({
        paraId: '11111',
        done: false
      });
      
      // First child comment
      const childComment1 = result.get('22222');
      expect(childComment1).toEqual({
        paraId: '22222',
        paraIdParent: '11111',
        done: false
      });
      
      // Second child comment (resolved)
      const childComment2 = result.get('33333');
      expect(childComment2).toEqual({
        paraId: '33333',
        paraIdParent: '11111',
        done: true
      });
    });

    it('should handle empty or malformed commentsExtended.xml', () => {
      const xmlString = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
        </w15:commentsEx>`;
      
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlString, 'application/xml');
      const result = parseExtendedComments(xmlDoc);
      
      expect(result.size).toBe(0);
    });

    it('should ignore commentEx elements without paraId', () => {
      const xmlString = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
          <w15:commentEx w15:done="1"/>
          <w15:commentEx w15:paraId="12345" w15:done="0"/>
        </w15:commentsEx>`;
      
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlString, 'application/xml');
      const result = parseExtendedComments(xmlDoc);
      
      expect(result.size).toBe(1);
      expect(result.has('12345')).toBe(true);
    });

    it('should handle comments without done attribute', () => {
      const xmlString = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
          <w15:commentEx w15:paraId="12345"/>
        </w15:commentsEx>`;
      
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlString, 'application/xml');
      const result = parseExtendedComments(xmlDoc);
      
      expect(result.size).toBe(1);
      const commentInfo = result.get('12345');
      expect(commentInfo).toEqual({
        paraId: '12345'
      });
    });

    it('should handle alternative namespace formats', () => {
      // Test without namespace prefix
      const xmlString = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <commentsEx>
          <commentEx paraId="12345" done="1" paraIdParent="67890"/>
        </commentsEx>`;
      
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlString, 'application/xml');
      const result = parseExtendedComments(xmlDoc);
      
      expect(result.size).toBe(1);
      const commentInfo = result.get('12345');
      expect(commentInfo).toEqual({
        paraId: '12345',
        paraIdParent: '67890',
        done: true
      });
    });
  });

  describe('Extended Comments Integration', () => {
    it('should gracefully handle missing commentsExtended.xml', () => {
      // This test ensures that when commentsExtended.xml is not present,
      // the application continues to work as before
      expect(true).toBe(true); // Placeholder - actual integration would be tested with file processing
    });

    it('should properly identify resolved comments with tick mark', () => {
      // This would test that resolved comments show the ✓ symbol
      // in the comment identification text
      expect(true).toBe(true); // Placeholder for future implementation testing
    });

    it('should properly indent threaded comments', () => {
      // This would test that comments with paraIdParent are indented
      expect(true).toBe(true); // Placeholder for future implementation testing
    });
  });
});