// Test script to verify our changes are working
import { Document, Paragraph, TextRun, Packer } from 'docx';
import fs from 'fs';

const doc = new Document({
  sections: [
    {
      properties: {},
      children: [
        new Paragraph({
          children: [
            new TextRun("This is a test document with annotations. "),
            new TextRun("This text will be highlighted for testing purposes.")
          ]
        }),
        new Paragraph({
          children: [
            new TextRun("Second paragraph for testing. [This is in square brackets]")
          ]
        })
      ]
    }
  ]
});

// Create the document and save it
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync('/tmp/test_document_generated.docx', buffer);
  console.log('Test document created at /tmp/test_document_generated.docx');
}).catch(err => {
  console.error('Error creating document:', err);
});