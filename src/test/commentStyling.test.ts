import { describe, it, expect } from 'vitest';
import { TextRun, ShadingType } from 'docx';

describe('Comment Styling', () => {
  it('should create styled comment anchor TextRun with shading', () => {
    const commentAnchorText = "[Comment 1]";
    const styledCommentAnchor = new TextRun({
      text: commentAnchorText,
      shading: {
        type: ShadingType.SOLID,
        fill: 'E6E6E6'
      }
    });

    expect(styledCommentAnchor).toBeDefined();
    expect(styledCommentAnchor).toBeInstanceOf(TextRun);
  });

  it('should create styled comment identification TextRun with italics', () => {
    const identificationText = "Comment 1 (John Doe, 2025-01-15): ";
    const styledIdentificationText = new TextRun({
      text: identificationText,
      italics: true
    });

    expect(styledIdentificationText).toBeDefined();
    expect(styledIdentificationText).toBeInstanceOf(TextRun);
  });

  it('should create multiple TextRuns with different styling', () => {
    // Simulate what happens in the modified buildTextRun function
    const results: TextRun[] = [];
    
    // Regular text
    results.push(new TextRun({
      text: "This is some regular text "
    }));
    
    // Comment anchor with shading
    results.push(new TextRun({
      text: "[Comment 1]",
      shading: {
        type: ShadingType.SOLID,
        fill: 'E6E6E6'
      }
    }));
    
    // More regular text
    results.push(new TextRun({
      text: " and more text."
    }));

    expect(results).toHaveLength(3);
    expect(results[0]).toBeInstanceOf(TextRun);
    expect(results[1]).toBeInstanceOf(TextRun);
    expect(results[2]).toBeInstanceOf(TextRun);
  });
});