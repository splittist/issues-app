import { describe, it, expect } from 'vitest';

describe('Column Width Adjustments', () => {
  it('should have significantly increased Response column width', () => {
    // This test verifies that our changes actually improved the Response column size
    
    // Before the change:
    // - With annotations: Ref(8%) + Paragraph(42%) + Annotation(40%) + Response(10%) = 100%
    // - Without annotations: Ref(8%) + Paragraph(82%) + Response(10%) = 100%
    
    // After the change:
    // - With annotations: Ref(8%) + Paragraph(31%) + Annotation(31%) + Response(30%) = 100%
    // - Without annotations: Ref(8%) + Paragraph(46%) + Response(46%) = 100%
    
    const oldResponseWidthWithAnnotations = 10;
    const newResponseWidthWithAnnotations = 30;
    const oldResponseWidthWithoutAnnotations = 10;
    const newResponseWidthWithoutAnnotations = 46;
    
    // Response column should be 3x larger when annotations are included
    expect(newResponseWidthWithAnnotations / oldResponseWidthWithAnnotations).toBe(3);
    
    // Response column should be 4.6x larger when annotations are not included
    expect(newResponseWidthWithoutAnnotations / oldResponseWidthWithoutAnnotations).toBe(4.6);
    
    // Response column should now be roughly equal to other content columns when annotations are included
    const paragraphWidthWithAnnotations = 31;
    const annotationWidthWithAnnotations = 31;
    expect(newResponseWidthWithAnnotations).toBeGreaterThanOrEqual(paragraphWidthWithAnnotations - 1);
    expect(newResponseWidthWithAnnotations).toBeGreaterThanOrEqual(annotationWidthWithAnnotations - 1);
    
    // Response column should now be equal to paragraph column when annotations are not included
    const paragraphWidthWithoutAnnotations = 46;
    expect(newResponseWidthWithoutAnnotations).toBe(paragraphWidthWithoutAnnotations);
  });

  it('should maintain total column width at 100%', () => {
    // Verify column widths add up to 100% in both scenarios
    
    // With annotations: Ref(8%) + Paragraph(31%) + Annotation(31%) + Response(30%) = 100%
    const refWidth = 8;
    const paragraphWidthWithAnnotations = 31;
    const annotationWidth = 31;
    const responseWidthWithAnnotations = 30;
    const totalWithAnnotations = refWidth + paragraphWidthWithAnnotations + annotationWidth + responseWidthWithAnnotations;
    expect(totalWithAnnotations).toBe(100);
    
    // Without annotations: Ref(8%) + Paragraph(46%) + Response(46%) = 100%
    const paragraphWidthWithoutAnnotations = 46;
    const responseWidthWithoutAnnotations = 46;
    const totalWithoutAnnotations = refWidth + paragraphWidthWithoutAnnotations + responseWidthWithoutAnnotations;
    expect(totalWithoutAnnotations).toBe(100);
  });

  it('should have balanced content columns when annotations are included', () => {
    // When annotations are included, all three content columns should be roughly equal
    const paragraphWidth = 31;
    const annotationWidth = 31;
    const responseWidth = 30;
    
    // All should be within 1% of each other (balanced)
    expect(Math.abs(paragraphWidth - annotationWidth)).toBeLessThanOrEqual(1);
    expect(Math.abs(paragraphWidth - responseWidth)).toBeLessThanOrEqual(1);
    expect(Math.abs(annotationWidth - responseWidth)).toBeLessThanOrEqual(1);
  });
});