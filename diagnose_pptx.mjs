import AdmZip from 'adm-zip';
import { readFileSync, writeFileSync } from 'fs';

const filePath = process.argv[2] || '/home/ubuntu/upload/converted.pptx';
console.log(`Inspecting: ${filePath}\n`);

const zip = new AdmZip(filePath);
const entries = zip.getEntries();

// List all files in the zip
console.log('=== ZIP CONTENTS ===');
for (const e of entries) {
  console.log(`  ${e.entryName} (${e.header.size} bytes)`);
}
console.log();

// Check [Content_Types].xml for duplicate entries
const ctEntry = zip.getEntry('[Content_Types].xml');
if (ctEntry) {
  const ct = ctEntry.getData().toString('utf-8');
  console.log('=== [Content_Types].xml ===');
  console.log(ct.substring(0, 2000));
  console.log();
  
  // Check for duplicate slide master entries
  const masterMatches = ct.match(/slideMaster/g) || [];
  const slideMatches = ct.match(/\/slides\/slide\d/g) || [];
  console.log(`  slideMaster references: ${masterMatches.length}`);
  console.log(`  slide references: ${slideMatches.length}`);
  console.log();
}

// Check slide XMLs for XML validity
const slideEntries = entries.filter(e => e.entryName.match(/^ppt\/slides\/slide\d+\.xml$/));
console.log(`=== SLIDE XML VALIDATION (${slideEntries.length} slides) ===`);

for (const entry of slideEntries) {
  const xml = entry.getData().toString('utf-8');
  
  // Check for unclosed tags
  const openTags = (xml.match(/<a:rPr[^/]/g) || []).length;
  const closeTags = (xml.match(/<\/a:rPr>/g) || []).length;
  const selfClosing = (xml.match(/<a:rPr[^>]*\/>/g) || []).length;
  
  // Check for malformed rPr (cut off mid-element)
  const badRpr = xml.match(/<a:rPr[^>]*>[^<]*\/>/g) || [];
  
  // Check for runs without closing tag
  const openRuns = (xml.match(/<a:r>/g) || []).length;
  const closeRuns = (xml.match(/<\/a:r>/g) || []).length;
  
  // Check for <a:t> without closing
  const openT = (xml.match(/<a:t[^>]*>/g) || []).length;
  const closeT = (xml.match(/<\/a:t>/g) || []).length;
  
  const issues = [];
  if (openTags !== closeTags + selfClosing) issues.push(`rPr mismatch: ${openTags} open, ${closeTags} close, ${selfClosing} self-closing`);
  if (openRuns !== closeRuns) issues.push(`<a:r> mismatch: ${openRuns} open vs ${closeRuns} close`);
  if (openT !== closeT) issues.push(`<a:t> mismatch: ${openT} open vs ${closeT} close`);
  if (badRpr.length > 0) issues.push(`Bad rPr patterns: ${badRpr.slice(0,3).join(', ')}`);
  
  if (issues.length > 0) {
    console.log(`  ❌ ${entry.entryName}: ${issues.join('; ')}`);
    // Show context around first issue
    const idx = xml.indexOf('<a:r>');
    if (idx > 0) {
      console.log(`     First <a:r> context: ...${xml.substring(idx, idx + 300)}...`);
    }
  } else {
    console.log(`  ✓ ${entry.entryName}: OK (${openRuns} runs, ${openTags + selfClosing} rPr)`);
  }
}

// Check ppt/presentation.xml for slide references
const presEntry = zip.getEntry('ppt/presentation.xml');
if (presEntry) {
  const pres = presEntry.getData().toString('utf-8');
  const slideRefs = pres.match(/r:id="rId\d+"/g) || [];
  console.log(`\n=== presentation.xml slide refs: ${slideRefs.length} ===`);
}

// Check _rels for slide relationships
const slideRelsEntries = entries.filter(e => e.entryName.match(/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/));
console.log(`\n=== SLIDE RELS (${slideRelsEntries.length} files) ===`);
for (const entry of slideRelsEntries.slice(0, 3)) {
  const xml = entry.getData().toString('utf-8');
  console.log(`  ${entry.entryName}:`);
  console.log(`    ${xml.substring(0, 400)}`);
}
