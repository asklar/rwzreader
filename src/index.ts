#!/usr/bin/env node

import * as fs from 'fs';
import { StreamBuffer } from './stream-buffer.js';
import { RulesFile } from './rwz-parser.js';

// Re-export the library API
export { StreamBuffer } from './stream-buffer.js';
export { OutlookRulesReadError, softAssert } from './errors.js';
export { OXCDATA, PropertyValueHeader, PropertyValueArray } from './oxcdata.js';
export { RuleElement } from './rule-element.js';
export {
  RulesFile, RulesHeader, RulesFooter, RuleHeader, Rule, OleDateTime,
  type OutlookRulesVersion,
} from './rwz-parser.js';
export * from './rule-element-data.js';

export function parseRwzFile(input: Buffer): RulesFile {
  return RulesFile.parse(new StreamBuffer(input));
}

// --- CLI entrypoint ---
function main() {
  const inputPath = process.argv[2];
  const outputPath = process.argv[3] || 'outlook-rules.json';

  if (!inputPath) {
    console.error('Usage: rwzreader <input.rwz> [output.json]');
    process.exit(1);
  }

  if (!fs.existsSync(inputPath)) {
    console.error(`File not found: ${inputPath}`);
    process.exit(1);
  }

  const content = fs.readFileSync(inputPath);
  const rf = parseRwzFile(content);
  fs.writeFileSync(outputPath, JSON.stringify(rf, null, 2));
  console.log(`Converted ${inputPath} -> ${outputPath}`);
}

main();
