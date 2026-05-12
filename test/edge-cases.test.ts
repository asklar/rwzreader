import { describe, it, expect } from 'vitest';
import { StreamBuffer } from '../src/stream-buffer.js';
import { RulesHeader, RulesFile } from '../src/rwz-parser.js';
import { RuleElement } from '../src/rule-element.js';
import { OutlookRulesReadError } from '../src/errors.js';
import { buildOutlook2019Header, buildMinimalRwz, u32 } from './helpers.js';

describe('edge cases', () => {
  it('truncated buffer throws during parsing', () => {
    // Build a header that claims 5 rules, but provide no rule data
    const buf = buildOutlook2019Header(5);
    const sb = new StreamBuffer(buf);
    expect(() => RulesFile.parse(sb)).toThrow();
  });

  it('unknown element ID throws with hex in message', () => {
    const payload = Buffer.concat([u32(0xbad), u32(0)]);
    const sb = new StreamBuffer(payload);
    try {
      RuleElement.parse(sb);
      expect.fail('should have thrown');
    } catch (e: any) {
      expect(e).toBeInstanceOf(OutlookRulesReadError);
      expect(e.message).toContain('0xbad');
    }
  });

  it('zero-rule file parses successfully', () => {
    const buf = buildMinimalRwz(0);
    const sb = new StreamBuffer(buf);
    const rf = RulesFile.parse(sb);
    expect(rf.rules).toHaveLength(0);
    expect(rf.header?.numberOfRules).toBe(0);
  });
});
