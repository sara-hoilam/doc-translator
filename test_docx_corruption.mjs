import fs from 'fs';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const AdmZip = require('adm-zip');

// Create a minimal valid DOCX file
const zip = new AdmZip();

// Add [Content_Types].xml
zip.addFile('[Content_Types].xml', Buffer.from(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>'
));

// Add _rels/.rels
zip.addFile('_rels/.rels', Buffer.from(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>'
));

// Add word/document.xml with test content
zip.addFile('word/document.xml', Buffer.from(
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:body>' +
  '</w:document>'
));

// Save the test DOCX
fs.writeFileSync('/tmp/test_original.docx', zip.toBuffer());
console.log('Created test DOCX file');

// Now test if we can read it back
const testZip = new AdmZip('/tmp/test_original.docx');
const docEntry = testZip.getEntry('word/document.xml');
if (docEntry) {
  console.log('✓ Can read word/document.xml from test DOCX');
  const xml = docEntry.getData().toString('utf-8');
  console.log('XML length:', xml.length);
  console.log('XML valid:', xml.includes('<w:document'));
} else {
  console.log('✗ Cannot find word/document.xml');
}
