import { describe, it, expect } from 'vitest';
import { StreamBuffer } from '../src/stream-buffer.js';
import {
  SimpleRuleElementData,
  ImportanceRuleElementData,
  SensitivityRuleElementData,
  SizeInSpecificRangeRuleElementData,
  DeferDeliveryRuleElementData,
  PathRuleElementData,
  CategoriesListRuleElementData,
  FlaggedForActionRuleElementData,
  StringsListRuleElementData,
  OnThisComputerOnlyRuleElementData,
} from '../src/rule-element-data.js';
import {
  buildSimpleElementPayload,
  buildExtReservedU32Payload,
  buildExtReservedStringPayload,
  u32, u8, u16, buildStringObject,
} from './helpers.js';

describe('SimpleRuleElementData', () => {
  it('reads extended=0', () => {
    const sb = new StreamBuffer(buildSimpleElementPayload());
    const data = new SimpleRuleElementData(sb);
    expect(data.extended).toBe(0);
    expect(sb.offset).toBe(4);
  });
});

describe('ImportanceRuleElementData', () => {
  it('reads importance value', () => {
    const sb = new StreamBuffer(buildExtReservedU32Payload(2));
    const data = new ImportanceRuleElementData(sb);
    expect(data.extended).toBe(1);
    expect(data.reserved).toBe(0);
    expect(data.importance).toBe(2);
  });
});

describe('SensitivityRuleElementData', () => {
  it('reads sensitivity value', () => {
    const sb = new StreamBuffer(buildExtReservedU32Payload(1));
    const data = new SensitivityRuleElementData(sb);
    expect(data.sensitivity).toBe(1);
  });
});

describe('SizeInSpecificRangeRuleElementData', () => {
  it('reads min and max sizes', () => {
    const payload = Buffer.concat([u32(1), u32(0), u32(100), u32(5000)]);
    const sb = new StreamBuffer(payload);
    const data = new SizeInSpecificRangeRuleElementData(sb);
    expect(data.minSize).toBe(100);
    expect(data.maxSize).toBe(5000);
  });
});

describe('DeferDeliveryRuleElementData', () => {
  it('reads minutes', () => {
    const sb = new StreamBuffer(buildExtReservedU32Payload(30));
    const data = new DeferDeliveryRuleElementData(sb);
    expect(data.minutes).toBe(30);
  });
});

describe('PathRuleElementData', () => {
  it('reads path string', () => {
    const sb = new StreamBuffer(buildExtReservedStringPayload('C:\\sounds\\ding.wav'));
    const data = new PathRuleElementData(sb);
    expect(data.path).toBe('C:\\sounds\\ding.wav');
  });
});

describe('CategoriesListRuleElementData', () => {
  it('reads categories string', () => {
    const sb = new StreamBuffer(buildExtReservedStringPayload('Important'));
    const data = new CategoriesListRuleElementData(sb);
    expect(data.categories).toBe('Important');
  });
});

describe('FlaggedForActionRuleElementData', () => {
  it('reads action name', () => {
    const sb = new StreamBuffer(buildExtReservedStringPayload('Follow up'));
    const data = new FlaggedForActionRuleElementData(sb);
    expect(data.actionName).toBe('Follow up');
  });
});

describe('StringsListRuleElementData', () => {
  it('reads multiple string entries', () => {
    // nEntries=2, then [flags=0 + stringObject] per entry
    const entry1 = Buffer.concat([u32(0), buildStringObject('hello')]);
    const entry2 = Buffer.concat([u32(0), buildStringObject('world')]);
    const payload = Buffer.concat([u32(2), entry1, entry2]);
    const sb = new StreamBuffer(payload);
    const data = new StringsListRuleElementData(sb);
    expect(data.entries).toHaveLength(2);
    expect(data.entries[0].value).toBe('hello');
    expect(data.entries[1].value).toBe('world');
  });

  it('reads zero entries', () => {
    const sb = new StreamBuffer(u32(0));
    const data = new StringsListRuleElementData(sb);
    expect(data.entries).toHaveLength(0);
  });
});

describe('OnThisComputerOnlyRuleElementData', () => {
  it('reads 16-byte uuid', () => {
    const uuid = Buffer.alloc(16, 0xab);
    const payload = Buffer.concat([u32(1), u32(0), uuid]);
    const sb = new StreamBuffer(payload);
    const data = new OnThisComputerOnlyRuleElementData(sb);
    expect(data.uuid.length).toBe(32); // 16 bytes as hex
  });
});
