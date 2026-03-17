/**
 * Verify the Content_Types fix on the existing converted.pptx
 */
import AdmZip from 'adm-zip';
import { readFileSync, writeFileSync } from 'fs';

const filePath = '/home/ubuntu/upload/converted.pptx';
const outPath = '/home/ubuntu/upload/converted_fixed.pptx';

console.log(`Reading: ${filePath}`);
const zip = new AdmZip(filePath);

const ctEntry = zip.getEntry('[Content_Types].xml');
if (!ctEntry) { console.error('No [Content_Types].xml found'); process.exit(1); }

let ctXml = ctEntry.getData().toString('utf-8');

// Count before
const beforeCount = (ctXml.match(/slideMaster/g) || []).length;
console.log(`Before: ${beforeCount} slideMaster references`);

// Apply the fix
let firstSeen = false;
ctXml = ctXml.replace(
  /<Override[^>]*\/ppt\/slideMasters\/slideMaster\d+\.xml[^>]*\/>/g,
  (match) => {
    if (!firstSeen) { firstSeen = true; return match; }
    return '';
  }
);

const afterCount = (ctXml.match(/slideMaster/g) || []).length;
console.log(`After: ${afterCount} slideMaster references`);

zip.updateFile('[Content_Types].xml', Buffer.from(ctXml, 'utf-8'));
zip.writeZip(outPath);
console.log(`\nFixed file written to: ${outPath}`);

// Verify the fixed file
const fixedZip = new AdmZip(outPath);
const fixedCt = fixedZip.getEntry('[Content_Types].xml').getData().toString('utf-8');
const fixedCount = (fixedCt.match(/slideMaster/g) || []).length;
console.log(`\nVerification: ${fixedCount} slideMaster references in fixed file`);

// Check that only slideMaster1.xml is referenced
const masterRefs = fixedCt.match(/slideMaster\d+\.xml/g) || [];
console.log(`Slide master files referenced: ${[...new Set(masterRefs)].join(', ')}`);

// Check all referenced files actually exist
const allEntries = new Set(fixedZip.getEntries().map(e => e.entryName));
const overrideRefs = [...fixedCt.matchAll(/PartName="([^"]+)"/g)].map(m => m[1].replace(/^\//, ''));
const missing = overrideRefs.filter(ref => !allEntries.has(ref));
if (missing.length > 0) {
  console.log(`\n❌ Missing files referenced in Content_Types: ${missing.join(', ')}`);
} else {
  console.log(`\n✅ All ${overrideRefs.length} referenced files exist in the zip`);
}
