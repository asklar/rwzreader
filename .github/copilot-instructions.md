# RWZ Reader

This is a TypeScript parser for Outlook `.rwz` (Rules Wizard) files. It converts binary `.rwz` files to JSON.

## Architecture
- Single-file TypeScript project: `src/index.ts`
- Binary format parser using a `StreamBuffer` class for reading LE integers, strings, etc.
- Classes model the RWZ format: `RulesHeader`, `RuleHeader`, `RuleElement`, `RulesFooter`
- Rule elements have typed data classes (e.g., `MoveToFolderRuleElementData`, `ImportanceRuleElementData`)
- Element IDs map to conditions (0xc8–0xf7), actions (0x12c–0x153), and exceptions (0x1f4–0x21b)

## Build & Run
- Build: `npm run build` (runs `tsc`, outputs to `bin/`)
- Run: `npm start` or `node bin/index.js <input.rwz> [output.json]`
- The project uses ESM modules (`"type": "module"` in package.json)

## Format Reference
- [OutlookRulesReader spec](https://github.com/hughbe/OutlookRulesReader)
- [MS-OXCDATA](https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/)
