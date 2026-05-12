import { describe, it, expect } from 'vitest';
import { StreamBuffer } from '../src/stream-buffer.js';
import { RulesHeader, RulesFile } from '../src/rwz-parser.js';
import { RuleElement } from '../src/rule-element.js';
import { OutlookRulesReadError } from '../src/errors.js';
import {
  buildOutlook2019Header, buildMinimalRwz,
  buildSimpleElementPayload, u32,
} from './helpers.js';

describe('RulesHeader.parse', () => {
  it('detects Outlook 2019 version', () => {
    const buf = buildOutlook2019Header(0);
    const sb = new StreamBuffer(buf);
    const header = RulesHeader.parse(sb);
    expect(header.version).toBe('outlook2019');
    expect(header.numberOfRules).toBe(0);
  });

  it('detects Outlook 2007 version', () => {
    const buf = buildOutlook2019Header(0);
    buf.writeUInt32LE(1200000, 0);    // outlook2007 signature
    buf.writeUInt32LE(0x06140000, 4); // valid flags for 2007
    const sb = new StreamBuffer(buf);
    const header = RulesHeader.parse(sb);
    expect(header.version).toBe('outlook2007');
  });

  it('reports correct number of rules', () => {
    const buf = buildOutlook2019Header(42);
    const sb = new StreamBuffer(buf);
    const header = RulesHeader.parse(sb);
    expect(header.numberOfRules).toBe(42);
  });
});

describe('RuleElement.parse', () => {
  it('parses SimpleRuleElementData for "sent only to me" (0xc9)', () => {
    const payload = Buffer.concat([u32(0xc9), buildSimpleElementPayload()]);
    const sb = new StreamBuffer(payload);
    const elem = RuleElement.parse(sb);
    expect(elem.id).toBe(0xc9);
    expect(elem.description).toBe('sent only to me');
  });

  it('parses "stop processing more rules" (0x142)', () => {
    const payload = Buffer.concat([u32(0x142), buildSimpleElementPayload()]);
    const sb = new StreamBuffer(payload);
    const elem = RuleElement.parse(sb);
    expect(elem.id).toBe(0x142);
    expect(elem.description).toBe('stop processing more rules');
  });

  it('parses exception "except where my name is in the To box" (0x1f4)', () => {
    const payload = Buffer.concat([u32(0x1f4), buildSimpleElementPayload()]);
    const sb = new StreamBuffer(payload);
    const elem = RuleElement.parse(sb);
    expect(elem.id).toBe(0x1f4);
    expect(elem.description).toContain('except');
  });

  it('throws on unknown element ID', () => {
    const payload = Buffer.concat([u32(0xffffff), buildSimpleElementPayload()]);
    const sb = new StreamBuffer(payload);
    expect(() => RuleElement.parse(sb)).toThrow(OutlookRulesReadError);
    expect(() => {
      sb.offset = 0;
      RuleElement.parse(sb);
    }).toThrow(/0xffffff/);
  });
});

describe('RulesFile.parse', () => {
  it('parses a minimal 1-rule RWZ', () => {
    const buf = buildMinimalRwz(1);
    const sb = new StreamBuffer(buf);
    const rf = RulesFile.parse(sb);
    expect(rf.header?.version).toBe('outlook2019');
    expect(rf.rules).toHaveLength(1);
    expect(rf.rules[0].elements).toHaveLength(2);
  });

  it('parses multiple rules', () => {
    const buf = buildMinimalRwz(3);
    const sb = new StreamBuffer(buf);
    const rf = RulesFile.parse(sb);
    expect(rf.rules).toHaveLength(3);
  });

  it('parses footer', () => {
    const buf = buildMinimalRwz(1);
    const sb = new StreamBuffer(buf);
    const rf = RulesFile.parse(sb);
    expect(rf.footer).toBeDefined();
  });

  it('round-trips to JSON', () => {
    const buf = buildMinimalRwz(2);
    const sb = new StreamBuffer(buf);
    const rf = RulesFile.parse(sb);
    const json = JSON.stringify(rf);
    const parsed = JSON.parse(json);
    expect(parsed.rules).toHaveLength(2);
    expect(parsed.header.version).toBe('outlook2019');
  });
});
