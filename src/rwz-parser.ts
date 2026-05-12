import { StreamBuffer } from './stream-buffer.js';
import { OutlookRulesReadError, softAssert } from './errors.js';
import { RuleElement } from './rule-element.js';

export type OutlookRulesVersion =
  | 'noSignature' | 'noSignatureOutlook2003'
  | 'outlook98' | 'outlook2000' | 'outlook2002'
  | 'outlook2003' | 'outlook2007' | 'outlook2019';

export class OleDateTime {
  public status?: 'Valid' | 'Null';
  public timestamp?: number;

  public get createdOn(): Date { return new Date(1900, 1, this.timestamp); }

  public static parse(sb: StreamBuffer) {
    const dt = new OleDateTime();
    const status = sb.readUInt32();
    dt.status = status === 0 ? 'Valid' : 'Null';
    dt.timestamp = sb.readDouble();
    return dt;
  }
}

export class RulesHeader {
  public version: OutlookRulesVersion = 'outlook2019';
  public signature?= 0x00140000;
  public flags?= 0x06140000;
  public numberOfRules = 0;
  public unknown = new Array<number | undefined>(9);

  private constructor() {}

  public static parse(buf: StreamBuffer) {
    const rh = new RulesHeader();
    const peekedSignature = buf.readUInt32();
    switch (peekedSignature) {
      case 1310720: rh.version = 'outlook2019'; break;
      case 1200000: rh.version = 'outlook2007'; break;
      case 1100000: rh.version = 'outlook2003'; break;
      case 1000000: rh.version = 'outlook2002'; break;
      case 980413: rh.version = 'outlook2000'; break;
      case 970812: rh.version = 'outlook98'; break;
      case 0: rh.version = 'noSignatureOutlook2003'; break;
      default: rh.version = 'noSignature'; break;
    }
    if (rh.version != 'noSignatureOutlook2003' && rh.version != 'noSignature') {
      rh.signature = peekedSignature;
      if (rh.version >= 'outlook2002') {
        try {
          const flags = buf.readUInt32();
          rh.flags = flags;
          if ((rh.version == 'outlook2019' && flags == 0x06140000) ||
            (rh.version == 'outlook2007' && flags == 0x06140000) ||
            (rh.version == 'outlook2007' && (flags == 0x06140000 || flags == 0x05124F80)) ||
            (rh.version == 'outlook2003' && flags == 0x04140000) ||
            (rh.version == 'outlook2002' && (flags == 0x03140000 || flags == 0x06140000 || flags == 0x03124F80 || flags == 0x05124F80))) {
          } else {
            throw new OutlookRulesReadError('invalidSignature');
          }
        } catch (e) { /* swallow for compatibility */ }
      } else { rh.flags = undefined; }
    } else { rh.signature = undefined; rh.flags = undefined; }
    if (rh.version !== 'noSignature') { rh.unknown[1] = buf.readUInt32(); if (rh.unknown[1] !== 0x00000000) throw new OutlookRulesReadError('corrupted 1'); } else { rh.unknown[1] = undefined; }
    if (rh.version !== 'noSignature') { rh.unknown[2] = buf.readUInt32(); if (rh.unknown[2] !== 0x00000000) throw new OutlookRulesReadError('corrupted 2'); } else { rh.unknown[2] = undefined; }
    if (rh.version != 'noSignature') { rh.unknown[3] = buf.readUInt32(); if (rh.unknown[3] != 0x00000000) throw new OutlookRulesReadError('corrupted 3'); } else { rh.unknown[3] = undefined; }
    if (rh.version != 'noSignature') { rh.unknown[4] = buf.readUInt32(); } else { rh.unknown[4] = undefined; }
    if (rh.version != 'noSignature') { rh.unknown[5] = buf.readUInt32(); } else { rh.unknown[6] = undefined; }
    if (rh.version != 'noSignature') { rh.unknown[6] = buf.readUInt32(); } else { rh.unknown[6] = undefined; }
    if (rh.version != 'noSignature') { rh.unknown[7] = buf.readUInt32(); } else { rh.unknown[7] = undefined; }
    if (rh.version != 'noSignature') { rh.unknown[8] = buf.readUInt32(); if (rh.unknown[8] != 0x00000001) throw new OutlookRulesReadError('corrupted 8'); } else { rh.unknown[9] = undefined; }
    if (rh.version >= 'outlook2002' || rh.version == 'noSignatureOutlook2003') { rh.unknown[9] = buf.readUInt32(); } else { rh.unknown[9] = undefined; }
    rh.numberOfRules = buf.readUInt16();
    const extra = buf.readUInt16();
    return rh;
  }
}

export class RulesFooter {
  public templateDirectoryLength?: number;
  public templateDirectory?: string;
  public creationDate?: OleDateTime;
  public unknown?: number;

  private constructor() {}

  public static parse(sb: StreamBuffer) {
    const rf = new RulesFooter();
    rf.templateDirectoryLength = sb.readUInt32();
    rf.templateDirectory = sb.readString(rf.templateDirectoryLength!);
    rf.creationDate = OleDateTime.parse(sb);
    rf.unknown = sb.readUInt32();
    return rf;
  }
}

export class RuleHeader {
  public signature?: number;
  public name?: string;
  public enabled?: boolean;
  public unknown = Array<number | undefined>(4);
  public dataSize?: number;
  public nRuleElements?: number;
  public padding?: number;
  public separator?: number;
  public classNameLength?: number;
  public className?: string;

  private constructor() {}

  public static parse(sb: StreamBuffer, index: number, totalRules: number) {
    const rh = new RuleHeader();
    rh.signature = sb.readUInt16();
    rh.name = sb.readStringObject();
    rh.enabled = sb.readUInt32() !== 0;
    for (let i = 0; i < 4; i++) { rh.unknown[i] = sb.readUInt32(); }
    rh.dataSize = sb.readUInt32();
    rh.nRuleElements = sb.readUInt16();
    rh.separator = sb.readUInt16();
    if (rh.separator === 0xffff) {
      softAssert(index === 0, `separator 0xFFFF expected only for first rule, got at index ${index}`);
      rh.padding = sb.readUInt16();
      rh.classNameLength = sb.readUInt16();
      softAssert(rh.classNameLength === 'CRuleElement'.length);
      rh.className = sb.readAsciiString(rh.classNameLength!);
      softAssert(rh.className === 'CRuleElement');
    } else if (rh.separator === 0x8001) {
      softAssert(index !== 0, `separator 0x8001 unexpected for first rule`);
    } else if (rh.separator === 0) {
      // tolerate
    } else {
      console.warn(`Warning: unexpected separator value 0x${rh.separator!.toString(16)} at rule index ${index}`);
    }
    return rh;
  }
}

export class Rule {
  public header?: RuleHeader;
  public elements: RuleElement[] = [];

  private constructor() {}

  public static parse(sb: StreamBuffer, index: number, totalRules: number) {
    const r = new Rule();
    r.header = RuleHeader.parse(sb, index, totalRules);
    for (let i = 0; i < r.header.nRuleElements!; i++) {
      const elem = RuleElement.parse(sb);
      r.elements.push(elem);
      if (i !== r.header.nRuleElements! - 1) {
        const separator = sb.readUInt16();
        softAssert(separator === 0x8001, `expected element separator 0x8001, got 0x${separator.toString(16)}`);
      }
    }
    return r;
  }
}

export class RulesFile {
  public header?: RulesHeader;
  public rules: Rule[] = [];
  public footer?: RulesFooter;

  private constructor() {}

  public static parse(buf: StreamBuffer) {
    const rf = new RulesFile();
    rf.header = RulesHeader.parse(buf);
    for (let i = 0; i < rf.header.numberOfRules; i++) {
      const rule = Rule.parse(buf, i, rf.header.numberOfRules);
      rf.rules.push(rule);
      if (i !== rf.header.numberOfRules - 1) {
        const separator = buf.readUInt16();
        softAssert(separator === 0, `expected inter-rule separator 0, got 0x${separator.toString(16)}`);
      }
    }
    try { rf.footer = RulesFooter.parse(buf); } catch (e) { console.warn(`Warning: could not parse footer: ${e}`); }
    return rf;
  }
}
