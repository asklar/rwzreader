# rwzreader

Parse Outlook Rules Wizard (`.rwz`) files and convert them to JSON.

## Quick Start

```bash
# No install needed — just run it
npx rwzreader input.rwz output.json
```

## Install

```bash
npm install -g rwzreader
rwzreader input.rwz output.json
```

## CLI Usage

```
rwzreader <input.rwz> [output.json]
```

- `input.rwz` — path to the Outlook `.rwz` rules file (required)
- `output.json` — path for JSON output (default: `outlook-rules.json`)

## Library API

```typescript
import { parseRwzFile } from 'rwzreader';
import * as fs from 'fs';

const buffer = fs.readFileSync('rules.rwz');
const rulesFile = parseRwzFile(buffer);

console.log(`Version: ${rulesFile.header?.version}`);
console.log(`Rules: ${rulesFile.rules.length}`);

for (const rule of rulesFile.rules) {
  console.log(`  ${rule.header?.name} (${rule.elements.length} elements)`);
}
```

### Exports

| Export | Description |
|---|---|
| `parseRwzFile(buf)` | Parse a Buffer into a `RulesFile` |
| `StreamBuffer` | Low-level binary reader |
| `RulesFile` | Parsed file (header + rules + footer) |
| `RulesHeader` | File header with version detection |
| `Rule` | Single rule (header + elements) |
| `RuleElement` | Condition, action, or exception element |

## Supported Versions

Outlook 98, 2000, 2002, 2003, 2007, and 2019.

## Supported Rule Elements

96 element types across three categories:

- **Conditions** — 30 types (name in To/CC, from sender, subject/body words, importance, sensitivity, categories, attachments, size range, date span, form type, RSS, etc.)
- **Actions** — 28 types (move/copy to folder, delete, forward, reply, flag, play sound, mark importance/sensitivity, Cc, redirect, print, run script, desktop alert, retention policy, etc.)
- **Exceptions** — 24 types (mirrors of conditions for exception handling)

Based on the [OutlookRulesReader specification](https://github.com/hughbe/OutlookRulesReader) by [@hughbe](https://github.com/hughbe).

## Project Structure

```
src/
  index.ts              # CLI entrypoint + library re-exports
  stream-buffer.ts      # Binary buffer reader (LE integers, strings)
  errors.ts             # OutlookRulesReadError + softAssert
  oxcdata.ts            # MS-OXCDATA property types + PropertyValueArray
  rule-element-data.ts  # 20+ data type classes for rule element payloads
  rule-element.ts       # Element ID → description + data class dispatch
  rwz-parser.ts         # RulesHeader, RulesFooter, RuleHeader, Rule, RulesFile
```

## Development

```bash
npm install
npm run build    # TypeScript → bin/
npm test         # vitest (42 tests)
```

## License

MIT
