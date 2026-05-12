import { StreamBuffer } from '../src/stream-buffer.js';

/** Write a UInt32LE into a buffer helper */
function u32(val: number): Buffer {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(val);
  return b;
}
function u16(val: number): Buffer {
  const b = Buffer.alloc(2);
  b.writeUInt16LE(val);
  return b;
}
function u8(val: number): Buffer {
  return Buffer.from([val]);
}

/**
 * Build a StringObject payload: length byte + UTF-16LE chars.
 * If length >= 0xFF, uses the 0xFF + uint16 + 2-byte pad prefix.
 */
export function buildStringObject(str: string): Buffer {
  const parts: Buffer[] = [];
  if (str.length < 0xff) {
    parts.push(u8(str.length));
  } else {
    parts.push(u8(0xff));
    parts.push(u16(str.length));
    parts.push(u16(0)); // 2-byte pad
  }
  for (const ch of str) {
    parts.push(u16(ch.charCodeAt(0)));
  }
  return Buffer.concat(parts);
}

/**
 * Build a SimpleRuleElementData payload: extended=0 (4 bytes).
 */
export function buildSimpleElementPayload(): Buffer {
  return u32(0);
}

/**
 * Build an extended+reserved+uint32 payload (e.g. ImportanceRuleElementData).
 */
export function buildExtReservedU32Payload(value: number): Buffer {
  return Buffer.concat([u32(1), u32(0), u32(value)]);
}

/**
 * Build an extended+reserved+stringObject payload (e.g. PathRuleElementData).
 */
export function buildExtReservedStringPayload(str: string): Buffer {
  return Buffer.concat([u32(1), u32(0), buildStringObject(str)]);
}

/**
 * Build a minimal Outlook 2019 RulesHeader for nRules rules.
 * Layout: signature(4) + flags(4) + unknown1-8(8×4) + unknown9(4) + nRules(2) + extra(2)
 */
export function buildOutlook2019Header(nRules: number): Buffer {
  const parts: Buffer[] = [];
  parts.push(u32(1310720));     // signature = outlook2019
  parts.push(u32(0x06140000));  // flags
  // unknown1-3: must be 0
  parts.push(u32(0)); parts.push(u32(0)); parts.push(u32(0));
  // unknown4-7: any value
  parts.push(u32(0)); parts.push(u32(0)); parts.push(u32(0)); parts.push(u32(0));
  // unknown8: must be 1
  parts.push(u32(1));
  // unknown9
  parts.push(u32(0));
  // numberOfRules
  parts.push(u16(nRules));
  // extra
  parts.push(u16(0));
  return Buffer.concat(parts);
}

/**
 * Build a RuleHeader for a given rule.
 * Layout: signature(2) + name(stringObject) + enabled(4) + unknown[0-3](4×4) +
 *         dataSize(4) + nRuleElements(2) + separator(2) [+ extras for first rule]
 */
export function buildRuleHeader(
  name: string,
  nElements: number,
  index: number,
  _totalRules: number,
): Buffer {
  const parts: Buffer[] = [];
  parts.push(u16(0x0001));                   // signature
  parts.push(buildStringObject(name));       // name
  parts.push(u32(1));                        // enabled = true
  for (let i = 0; i < 4; i++) parts.push(u32(0)); // unknown[0-3]
  parts.push(u32(0));                        // dataSize (we don't validate this)
  parts.push(u16(nElements));                // nRuleElements

  if (index === 0) {
    parts.push(u16(0xffff));                 // separator for first rule
    parts.push(u16(0));                      // padding
    const className = 'CRuleElement';
    parts.push(u16(className.length));
    parts.push(Buffer.from(className, 'ascii'));
  } else {
    parts.push(u16(0x8001));                 // separator for subsequent rules
  }
  return Buffer.concat(parts);
}

/**
 * Build a RuleElement: elementId(4) + data payload.
 */
export function buildRuleElement(elementId: number, dataPayload: Buffer): Buffer {
  return Buffer.concat([u32(elementId), dataPayload]);
}

/** Element separator between elements within a rule */
export function buildElementSeparator(): Buffer {
  return u16(0x8001);
}

/** Inter-rule separator */
export function buildRuleSeparator(): Buffer {
  return u16(0);
}

/**
 * Build a minimal RulesFooter.
 * Layout: templateDirLength(4) + templateDir(UTF-16LE) + oleDateTime(4+8) + unknown(4)
 */
export function buildFooter(): Buffer {
  const parts: Buffer[] = [];
  parts.push(u32(0));        // templateDirectoryLength = 0
  // OleDateTime: status(4) + timestamp(8)
  parts.push(u32(0));        // status = Valid
  const ts = Buffer.alloc(8);
  ts.writeDoubleLE(0);
  parts.push(ts);
  parts.push(u32(0));        // unknown
  return Buffer.concat(parts);
}

/**
 * Build a complete minimal .rwz buffer with the given number of simple rules.
 * Each rule has 2 elements: Unknown (0x64) + ApplyRule (0x190).
 */
export function buildMinimalRwz(nRules: number): Buffer {
  const parts: Buffer[] = [];
  parts.push(buildOutlook2019Header(nRules));

  for (let i = 0; i < nRules; i++) {
    // Rule header
    parts.push(buildRuleHeader(`Rule ${i}`, 2, i, nRules));

    // Element 0x64 (Unknown): extended=1, reserved=0, flags=1
    const elem64 = Buffer.concat([u32(1), u32(0), u32(1)]);
    parts.push(buildRuleElement(0x64, elem64));

    // Element separator
    parts.push(buildElementSeparator());

    // Element 0x190 (ApplyRule): extended=1, reserved=0, flags=1 (after message arrives)
    const elem190 = Buffer.concat([u32(1), u32(0), u32(1)]);
    parts.push(buildRuleElement(0x190, elem190));

    // Inter-rule separator (except after last rule)
    if (i !== nRules - 1) {
      parts.push(buildRuleSeparator());
    }
  }

  parts.push(buildFooter());
  return Buffer.concat(parts);
}

export { u8, u16, u32 };
