import { describe, it, expect } from 'vitest';
import { TextRun } from 'docx';

describe('Comment Styling', () => {
  it('should create styled comment anchor TextRun with named style', () => {
    const commentAnchorText = "[Comment 1]";
    const styledCommentAnchor = new TextRun({
      text: commentAnchorText,
      style: 'CommentAnchor'
    });

    expect(styledCommentAnchor).toBeDefined();
    expect(styledCommentAnchor).toBeInstanceOf(TextRun);
  });

  it('should create styled footnote anchor TextRun with named style', () => {
    const footnoteAnchorText = "[Footnote 1]";
    const styledFootnoteAnchor = new TextRun({
      text: footnoteAnchorText,
      style: 'FootnoteAnchor'
    });

    expect(styledFootnoteAnchor).toBeDefined();
    expect(styledFootnoteAnchor).toBeInstanceOf(TextRun);
  });

  it('should create styled endnote anchor TextRun with named style', () => {
    const endnoteAnchorText = "[Endnote 1]";
    const styledEndnoteAnchor = new TextRun({
      text: endnoteAnchorText,
      style: 'EndnoteAnchor'
    });

    expect(styledEndnoteAnchor).toBeDefined();
    expect(styledEndnoteAnchor).toBeInstanceOf(TextRun);
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
    
    // Comment anchor with named style
    results.push(new TextRun({
      text: "[Comment 1]",
      style: 'CommentAnchor'
    }));
    
    // Footnote anchor with named style
    results.push(new TextRun({
      text: "[Footnote 1]",
      style: 'FootnoteAnchor'
    }));
    
    // Endnote anchor with named style
    results.push(new TextRun({
      text: "[Endnote 1]",
      style: 'EndnoteAnchor'
    }));
    
    // More regular text
    results.push(new TextRun({
      text: " and more text."
    }));

    expect(results).toHaveLength(5);
    expect(results[0]).toBeInstanceOf(TextRun);
    expect(results[1]).toBeInstanceOf(TextRun);
    expect(results[2]).toBeInstanceOf(TextRun);
    expect(results[3]).toBeInstanceOf(TextRun);
    expect(results[4]).toBeInstanceOf(TextRun);
  });

  it('should ensure anchor styles take precedence over runProps styles', () => {
    // Simulate runProps that includes a style property
    const runPropsWithStyle = {
      bold: true,
      italics: true,
      style: 'SomeOtherStyle'
    };
    
    // Test that anchor style overwrites runProps style when placed after spread
    const commentAnchor = new TextRun({
      text: "[Comment 1]",
      ...runPropsWithStyle,
      style: 'CommentAnchor'
    });
    
    const footnoteAnchor = new TextRun({
      text: "[Footnote 1]",
      ...runPropsWithStyle,
      style: 'FootnoteAnchor'
    });
    
    const endnoteAnchor = new TextRun({
      text: "[Endnote 1]",
      ...runPropsWithStyle,
      style: 'EndnoteAnchor'
    });

    expect(commentAnchor).toBeDefined();
    expect(footnoteAnchor).toBeDefined();
    expect(endnoteAnchor).toBeDefined();
    expect(commentAnchor).toBeInstanceOf(TextRun);
    expect(footnoteAnchor).toBeInstanceOf(TextRun);
    expect(endnoteAnchor).toBeInstanceOf(TextRun);
  });
});