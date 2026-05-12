import * as fs from 'fs';
import { Readable } from 'stream';
import { promisify } from 'util';

function softAssert(condition: boolean, msg?: string) {
  if (!condition) {
    console.warn(`Warning: assertion failed${msg ? ': ' + msg : ''}`);
  }
}

class StreamBuffer {
  public readUInt64() {
    const v = this.buffer.readBigUInt64LE(this.offset);
    this.offset += 8;
    return v;
  }
  public offset = 0;
  public constructor(private readonly buffer: Buffer) { }
  public readUInt32() {
    const v = this.buffer.readUInt32LE(this.offset);
    this.offset += 4;
    return v;
  }
  public readUInt16() {
    const v = this.buffer.readUInt16LE(this.offset);
    this.offset += 2;
    return v;
  }
  public readDouble() {
    const v = this.buffer.readDoubleLE(this.offset);
    this.offset += 8;
    return v;
  }
  public readString(len: number) {
    const b = Buffer.alloc(len);
    for (let i = 0; i < len; i++) {
      b[i] = this.readUInt16();
    }
    return b.toString();
  }
  public readStringObject() {
    let length = this.readUInt8();
    if (length === 0xff) {
      length = this.readUInt16();
      this.offset += 2;
    }
    const str = this.readString(length);
    return str;
  }
  public readUInt8() {
    const v = this.buffer.readUInt8(this.offset);
    this.offset++;
    return v;
  }
  public readAsciiString(len: number) {
    const b = Buffer.alloc(len);
    try {
      for (let i = 0; i < len; i++) {
        b[i] = this.readUInt8();
      }
      return b.toString();
    } catch (e) {
      return '';
    }
  }
  public readStringUntilNullTerminator() {
    const v = Buffer.alloc(260);
    try {
      let i = 0;
      for (let c = this.readUInt16(); c !== 0; c = this.readUInt16()) {
        v.writeUInt16LE(c, i++);
      }
      return v.toString('utf8', 0, i);
    } catch (e) {
      return '';
    }
  }
  public readAsciiUntilNullTerminator() {
    const v = Buffer.alloc(260);
    let i = 0;
    for (let c = this.readUInt8(); c !== 0; c = this.readUInt8()) {
      v.writeUInt16LE(c, i++);
    }
    return v.toString('utf8', 0, i);
  }
  public readBytes(len: number): Buffer {
    const b = this.buffer.subarray(this.offset, this.offset + len);
    this.offset += len;
    return Buffer.from(b as any);
  }
  public get remaining() {
    return this.buffer.length - this.offset;
  }
}

type OutlookRulesVersion = 'noSignature' | 'noSignatureOutlook2003' | 'outlook98' | 'outlook2000' | 'outlook2002' | 'outlook2003' | 'outlook2007' | 'outlook2019';

class OutlookRulesReadError extends Error { }

class RulesHeader {
  public version: OutlookRulesVersion = 'outlook2019';
  public signature?= 0x00140000;
  public flags?= 0x06140000;
  public numberOfRules = 0;
  public unknown = new Array<number | undefined>(9);
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
          console.log(`Flags: ${flags}`);
          if ((rh.version == 'outlook2019' && flags == 0x06140000) ||
            (rh.version == 'outlook2007' && flags == 0x06140000) ||
            (rh.version == 'outlook2007' && (flags == 0x06140000 || flags == 0x05124F80)) ||
            (rh.version == 'outlook2003' && flags == 0x04140000) ||
            (rh.version == 'outlook2002' && (flags == 0x03140000 || flags == 0x06140000 || flags == 0x03124F80 || flags == 0x05124F80))) {
          } else {
            throw new OutlookRulesReadError('invalidSignature');
          }
        } catch (e) { console.log(e); }
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
  private constructor() { }
}

class OleDateTime {
  private status?: 'Valid' | 'Null';
  private timestamp?: number;
  public get createdOn(): Date { return new Date(1900, 1, this.timestamp); }
  private OleDateTime() {}
  public static parse(sb: StreamBuffer) {
    const dt = new OleDateTime();
    const status = sb.readUInt32();
    dt.status = status === 0 ? 'Valid' : 'Null';
    dt.timestamp = sb.readDouble();
    return dt;
  }
}

class RulesFooter {
  private templateDirectoryLength?: number;
  private templateDirectory?: string;
  private creationDate?: OleDateTime;
  private unknown?: number;
  private constructor() {}
  public static parse(sb: StreamBuffer) {
    const rf = new RulesFooter();
    rf.templateDirectoryLength = sb.readUInt32();
    rf.templateDirectory = sb.readString(rf.templateDirectoryLength);
    rf.creationDate = OleDateTime.parse(sb);
    rf.unknown = sb.readUInt32();
    return rf;
  }
}

class RuleHeader {
  private signature?: number;
  private name?: string;
  private enabled?: boolean;
  private unknown = Array<number | undefined>(4);
  private dataSize?: number;
  public nRuleElements?: number;
  private padding?: number;
  private separator?: number;
  private classNameLength?: number;
  private className?: string;
  private constructor() {}
  public static parse(sb: StreamBuffer, index: number, totalRules: number) {
    const rh = new RuleHeader();
    rh.signature = sb.readUInt16();
    rh.name = sb.readStringObject();
    rh.enabled = sb.readUInt32() === 0 ? false : true;
    for (let i = 0; i < 4; i++) { rh.unknown[i] = sb.readUInt32(); }
    rh.dataSize = sb.readUInt32();
    rh.nRuleElements = sb.readUInt16();
    // Last rule has NO separator
    if (index === totalRules - 1) {
      // No separator for the last rule
    } else {
      rh.separator = sb.readUInt16();
      if (rh.separator === 0xffff) {
        softAssert(index === 0, `separator 0xFFFF expected only for first rule, got at index ${index}`);
        rh.padding = sb.readUInt16();
        rh.classNameLength = sb.readUInt16();
        softAssert(rh.classNameLength === 'CRuleElement'.length);
        rh.className = sb.readAsciiString(rh.classNameLength);
        softAssert(rh.className === 'CRuleElement');
      } else if (rh.separator === 0x8001) {
        softAssert(index !== 0, `separator 0x8001 unexpected for first rule`);
      } else if (rh.separator === 0) {
        // tolerate
      } else {
        console.warn(`Warning: unexpected separator value 0x${rh.separator.toString(16)} at rule index ${index}`);
      }
    }
    return rh;
  }
}

class RuleElementData {
  public constructor(sb: StreamBuffer) {}
}

class UnknownRuleElement0x64Data extends RuleElementData {
  public extended: number;
  public reserved: number;
  public flags: number;
  public constructor(sb: StreamBuffer) {
    super(sb);
    this.extended = sb.readUInt32();
    softAssert(this.extended === 0x1);
    this.reserved = sb.readUInt32();
    this.flags = sb.readUInt32();
  }
}

class ApplyRuleElementData extends UnknownRuleElement0x64Data {
  public when() {
    switch (this.flags) {
      case 0x1: return 'after the message arrives';
      case 0x4: return 'after I send the message';
      case 0x8: return 'after the server receives the message';
      default: throw new OutlookRulesReadError('unknown flag');
    }
  }
}

class SimpleRuleElementData extends RuleElementData {
  public extended?: number;
  public constructor(sb: StreamBuffer) {
    super(sb);
    this.extended = sb.readUInt32();
    softAssert(this.extended === 0);
  }
}

enum OXCDATA {
  PtypInteger16 = 2, PtypInteger32 = 3, PtypFloating32 = 4, PtypFloating64 = 5,
  PtypCurrency = 6, PtypFloatingTime = 7, PtypErrorCode = 0xa, PtypBoolean = 0xb,
  PtypInteger64 = 0x14, PtypString = 0x1f, PtypString8 = 0x1e, PtypTime = 0x40,
  PtypGuid = 0x48, PtypServerId = 0xfb, PtypRestriction = 0xfd, PtypRuleAction = 0xfe,
  PtypBinary = 0x102, PtypMultipleInteger16 = 0x1002, PtypMultipleInteger32 = 0x1003,
  PtypMultipleFloating32 = 0x1004, PtypMultipleFloating64 = 0x1005,
  PtypMultipleCurrency = 0x1006, PtypMultipleFloatingTime = 0x1007,
  PtypMultipleInteger64 = 0x1014, PtypMultipleString = 0x101f,
  PtypMultipleString8 = 0x101e, PtypMultipleTime = 0x1040,
  PtypMultipleGuid = 0x1048, PtypMultipleBinary = 0x1102,
  PtypUnspecified = 0, PtypNull = 1, PtypObject = 0xd, PtypEmbeddedTable = 0xd,
}

class PropertyValueHeader {
  public readonly dataType: OXCDATA;
  public readonly id: number;
  public readonly data: number[] = [];
  public type() { return OXCDATA[this.dataType]; }
  public constructor(sb: StreamBuffer) {
    this.dataType = sb.readUInt16();
    this.id = sb.readUInt16();
    this.data.push(sb.readUInt32());
    this.data.push(sb.readUInt32());
    this.data.push(sb.readUInt32());
  }
}

class PropertyValueArray {
  public propertyHeaders: PropertyValueHeader[] = [];
  public readonly properties: Record<number, any> = {};
  public constructor(sb: StreamBuffer) {
    const unknown = sb.readUInt32();
    const nProps = sb.readUInt32();
    const propDataSize = sb.readUInt32();
    const startPos = sb.offset;
    const endPosition = sb.offset + propDataSize;
    for (let i = 0; i < nProps; i++) {
      const ph = new PropertyValueHeader(sb);
      this.propertyHeaders.push(ph);
      const position = sb.offset;
      let value: any = undefined;
      switch (ph.dataType) {
        case OXCDATA.PtypInteger32: value = ph.data[1]; break;
        case OXCDATA.PtypErrorCode: value = ph.data[1]; break;
        case OXCDATA.PtypString: {
          const offset = ph.data[1];
          softAssert(offset >= 0); softAssert(offset <= endPosition);
          sb.offset = startPos + offset;
          value = sb.readStringUntilNullTerminator();
          softAssert(sb.offset <= endPosition);
          sb.offset = position;
          break;
        }
        case OXCDATA.PtypString8: {
          const offset = ph.data[1];
          softAssert(offset >= 0); softAssert(offset <= endPosition);
          sb.offset = startPos + offset;
          value = sb.readAsciiUntilNullTerminator();
          softAssert(sb.offset <= endPosition);
          sb.offset = position;
          break;
        }
        case OXCDATA.PtypBinary: {
          const offset = ph.data[1];
          softAssert(offset >= 0); softAssert(offset <= endPosition);
          sb.offset = startPos + offset;
          const length = ph.data[2];
          softAssert(length >= 0); softAssert(sb.offset + length <= endPosition);
          value = sb.readAsciiString(length);
          sb.offset = position;
          break;
        }
        case OXCDATA.PtypBoolean: { value = ph.data[1] !== 0; break; }
        default: { throw new Error('NYI'); }
      }
      this.properties[ph.id] = value;
    }
    sb.offset = endPosition;
  }
}

class PeopleOrPublicGroupListRuleElementData extends RuleElementData {
  public extended?: number;
  public reserved?: number;
  public values: PropertyValueArray[] = [];
  public constructor(sb: StreamBuffer) {
    super(sb);
    this.extended = sb.readUInt32();
    softAssert(this.extended === 1);
    this.reserved = sb.readUInt32();
    softAssert(this.reserved === 0);
    const nValues = sb.readUInt32();
    for (let i = 0; i < nValues; i++) { this.values.push(new PropertyValueArray(sb)); }
    const unknown1 = sb.readUInt32();
    softAssert(unknown1 === 1);
    const unknown2 = sb.readUInt32();
    softAssert(unknown2 === 0);
  }
}

class SearchEntry {
  public flags?: number;
  public value?: string;
  public constructor(sb: StreamBuffer) {
    this.flags = sb.readUInt32();
    softAssert(this.flags === 0);
    this.value = sb.readStringObject();
  }
}

class StringsListRuleElementData extends RuleElementData {
  public entries: SearchEntry[] = [];
  public constructor(sb: StreamBuffer) {
    super(sb);
    const nEntries = sb.readUInt32();
    for (let i = 0; i < nEntries; i++) { this.entries.push(new SearchEntry(sb)); }
  }
}

class FlaggedForActionRuleElementData extends RuleElementData {
  public extended?: number;
  public reserved?: number;
  public actionName?: string;
  public constructor(sb: StreamBuffer) {
    super(sb);
    this.extended = sb.readUInt32();
    softAssert(this.extended === 1);
    this.reserved = sb.readUInt32();
    softAssert(this.reserved === 0);
    this.actionName = sb.readStringObject();
  }
}

class FlatEntry {
  public readonly size: number;
  public constructor(sb: StreamBuffer) { this.size = sb.readUInt32(); }
}

class FolderEntryId extends FlatEntry {
  public readonly flags?: number;
  public readonly providerUID?: string;
  public readonly folderType?: number;
  public readonly databaseGuid?: string;
  public readonly globalCounter?: string;
  public constructor(sb: StreamBuffer) {
    const pos = sb.offset;
    super(sb);
    if (this.size !== 0) {
      this.flags = sb.readUInt32();
      softAssert(this.flags === 0);
      this.providerUID = sb.readAsciiString(16);
      this.folderType = sb.readUInt16();
      this.databaseGuid = sb.readAsciiString(16);
      this.globalCounter = sb.readUInt64().toString();
    }
    sb.offset = pos + this.size + 4;
  }
}

class StoreEntryId extends FlatEntry {
  public readonly flags: number;
  public readonly providerUID: string;
  public readonly version: number;
  public readonly flag: number;
  public readonly dllFileName: string;
  public readonly wrappedFlags: number;
  public readonly wrappedProvider: string;
  public readonly wrappedType: number;
  public readonly serverShortName: string;
  public readonly mailboxDN: string;
  public constructor(sb: StreamBuffer) {
    const pos = sb.offset;
    super(sb);
    this.flags = sb.readUInt32();
    softAssert(this.flags === 0);
    this.providerUID = sb.readAsciiString(16);
    this.version = sb.readUInt8();
    softAssert(this.version === 0);
    this.flag = sb.readUInt8();
    softAssert(this.flags === 0);
    this.dllFileName = sb.readAsciiString(14);
    softAssert(this.dllFileName === 'EMSMDB.DLL\0\0\0\0');
    this.wrappedFlags = sb.readUInt32();
    softAssert(this.wrappedFlags === 0);
    this.wrappedProvider = sb.readAsciiString(16);
    this.wrappedType = sb.readUInt32();
    softAssert(this.wrappedType === 0xc);
    this.serverShortName = sb.readAsciiUntilNullTerminator();
    this.mailboxDN = sb.readAsciiUntilNullTerminator();
    sb.offset = pos + this.size + 4;
  }
}

class MoveToFolderRuleElementData extends RuleElementData {
  public readonly extended: number;
  public readonly reserved: number;
  public readonly folderEntryId: FlatEntry;
  public readonly storeEntryId: FlatEntry;
  public readonly folderName: string;
  public readonly secondaryUserStore: boolean;
  public constructor(sb: StreamBuffer) {
    super(sb);
    this.extended = sb.readUInt32();
    this.reserved = sb.readUInt32();
    this.folderEntryId = new FolderEntryId(sb);
    this.storeEntryId = new StoreEntryId(sb);
    this.folderName = sb.readStringObject();
    this.secondaryUserStore = sb.readUInt32() !== 0;
  }
}

// --- New data types ---
class ImportanceRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public importance: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.importance = sb.readUInt32();
  }
}
class SensitivityRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public sensitivity: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.sensitivity = sb.readUInt32();
  }
}
class CategoriesListRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public categories: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.categories = sb.readStringObject();
  }
}
class PathRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public path: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.path = sb.readStringObject();
  }
}
class DisplayMessageInNewItemAlertWindowRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public message: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.message = sb.readStringObject();
  }
}
class FlagRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public days: number;
  public actionName: string; public unknown: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.days = sb.readUInt32(); this.actionName = sb.readStringObject();
    this.unknown = sb.readUInt32();
  }
}
class DeferDeliveryRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public minutes: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.minutes = sb.readUInt32();
  }
}
class PerformCustomActionRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public location: string; public name: string; public options: string; public actionValue: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.location = sb.readStringObject(); this.name = sb.readStringObject();
    this.options = sb.readStringObject(); this.actionValue = sb.readStringObject();
  }
}
class AutomaticReplyRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public messageEntryId: FlatEntry; public name: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.messageEntryId = new FlatEntry(sb);
    if (this.messageEntryId.size > 0) { sb.readBytes(this.messageEntryId.size); }
    this.name = sb.readStringObject();
  }
}
class RunScriptRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public name: string; public functionName: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.name = sb.readStringObject(); this.functionName = sb.readStringObject();
  }
}
class FlagForFollowUpRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public followUp: number; public actionName: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.followUp = sb.readUInt32(); this.actionName = sb.readStringObject();
  }
}
class ApplyRetentionPolicyRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public followUp: number; public guid: string; public name: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.followUp = sb.readUInt32(); this.guid = sb.readAsciiString(16);
    this.name = sb.readStringObject();
  }
}
class OnThisComputerOnlyRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public uuid: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.uuid = sb.readAsciiString(16);
  }
}
class WithSelectedPropertiesOfDocumentOrFormsRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
  }
}
class SizeInSpecificRangeRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public minSize: number; public maxSize: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.minSize = sb.readUInt32(); this.maxSize = sb.readUInt32();
  }
}
class ReceivedInSpecificDateSpanRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public startDate: OleDateTime; public endDate: OleDateTime;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.startDate = OleDateTime.parse(sb); this.endDate = OleDateTime.parse(sb);
  }
}
class FormTypeRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public formClass: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.formClass = sb.readStringObject();
  }
}
class ThroughAccountRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public accountName: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.accountName = sb.readStringObject();
  }
}
class SenderInSpecifiedAddressBookRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public addressBookName: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.addressBookName = sb.readStringObject();
  }
}

class RuleElement {
  public id?: number;
  public description?: string;
  public data?: RuleElementData;
  private constructor() {}
  public static parse(sb: StreamBuffer) {
    const re = new RuleElement();
    re.id = sb.readUInt32();
    switch (re.id) {
      // === Mandatory ===
      case 0x64: re.description = 'Unknown'; re.data = new UnknownRuleElement0x64Data(sb); break;
      case 0x190: re.description = 'type of message to which this rule applies'; re.data = new ApplyRuleElementData(sb); break;
      // === CONDITIONS (0xC8-0xF7) ===
      case 0xc8: re.description = 'where my name is in the To box'; re.data = new SimpleRuleElementData(sb); break;
      case 0xc9: re.description = 'sent only to me'; re.data = new SimpleRuleElementData(sb); break;
      case 0xca: re.description = 'where my name is not in the To box'; re.data = new SimpleRuleElementData(sb); break;
      case 0xcb: re.description = 'from <people or public group>'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0xcc: re.description = 'sent to <people or public group>'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0xcd: re.description = 'with specific words in the subject'; re.data = new StringsListRuleElementData(sb); break;
      case 0xce: re.description = 'with specific words in the body'; re.data = new StringsListRuleElementData(sb); break;
      case 0xcf: re.description = 'with specific words in the subject or body'; re.data = new StringsListRuleElementData(sb); break;
      case 0xd0: re.description = 'flagged for <action>'; re.data = new FlaggedForActionRuleElementData(sb); break;
      case 0xd2: re.description = 'marked as importance'; re.data = new ImportanceRuleElementData(sb); break;
      case 0xd3: re.description = 'marked as sensitivity'; re.data = new SensitivityRuleElementData(sb); break;
      case 0xd7: re.description = 'assigned to <category> category'; re.data = new CategoriesListRuleElementData(sb); break;
      case 0xdc: re.description = 'which is an automatic reply'; re.data = new SimpleRuleElementData(sb); break;
      case 0xde: re.description = 'which has an attachment'; re.data = new SimpleRuleElementData(sb); break;
      case 0xdf: re.description = 'with selected properties of documents or forms'; re.data = new WithSelectedPropertiesOfDocumentOrFormsRuleElementData(sb); break;
      case 0xe0: re.description = 'with a size in a specific range'; re.data = new SizeInSpecificRangeRuleElementData(sb); break;
      case 0xe1: re.description = 'received in a specific date span'; re.data = new ReceivedInSpecificDateSpanRuleElementData(sb); break;
      case 0xe2: re.description = 'where my name is in the CC box'; re.data = new SimpleRuleElementData(sb); break;
      case 0xe3: re.description = 'where my name is in the To or CC box'; re.data = new SimpleRuleElementData(sb); break;
      case 0xe4: re.description = 'uses the <form> form'; re.data = new FormTypeRuleElementData(sb); break;
      case 0xe5: re.description = 'with specific words in the sender address'; re.data = new StringsListRuleElementData(sb); break;
      case 0xe6: re.description = 'with specific words in the recipient address'; re.data = new StringsListRuleElementData(sb); break;
      case 0xe7: re.description = 'with specific words in the message header'; re.data = new StringsListRuleElementData(sb); break;
      case 0xe8: re.description = 'with specific words in the message header (alt)'; re.data = new StringsListRuleElementData(sb); break;
      case 0xea: re.description = 'on this computer only'; re.data = new OnThisComputerOnlyRuleElementData(sb); break;
      case 0xec: re.description = 'through the specified account'; re.data = new ThroughAccountRuleElementData(sb); break;
      case 0xed: re.description = 'sender is in specified address book'; re.data = new SenderInSpecifiedAddressBookRuleElementData(sb); break;
      case 0xf2: re.description = 'uses the <form> form (alt)'; re.data = new FormTypeRuleElementData(sb); break;
      case 0xf6: re.description = 'assigned to any category'; re.data = new SimpleRuleElementData(sb); break;
      case 0xf7: re.description = 'which is a meeting invitation or update'; re.data = new SimpleRuleElementData(sb); break;
      // === ACTIONS (0x12C-0x153) ===
      case 0x12c: re.description = 'move it to the specified folder'; re.data = new MoveToFolderRuleElementData(sb); break;
      case 0x12d: re.description = 'delete it'; re.data = new SimpleRuleElementData(sb); break;
      case 0x12e: re.description = 'forward it to people or public group'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x12f: re.description = 'reply using template'; re.data = new PathRuleElementData(sb); break;
      case 0x130: re.description = 'display a specific message in New Item Alert window'; re.data = new DisplayMessageInNewItemAlertWindowRuleElementData(sb); break;
      case 0x131: re.description = 'flag message for action'; re.data = new FlagRuleElementData(sb); break;
      case 0x132: re.description = 'clear the Message flag'; re.data = new SimpleRuleElementData(sb); break;
      case 0x133: re.description = 'assign it to category'; re.data = new CategoriesListRuleElementData(sb); break;
      case 0x136: re.description = 'play sound'; re.data = new PathRuleElementData(sb); break;
      case 0x137: re.description = 'mark it as importance'; re.data = new ImportanceRuleElementData(sb); break;
      case 0x138: re.description = 'mark it as sensitivity'; re.data = new SensitivityRuleElementData(sb); break;
      case 0x139: re.description = 'move a copy to the specified folder'; re.data = new MoveToFolderRuleElementData(sb); break;
      case 0x13a: re.description = 'notify me when it is read'; re.data = new SimpleRuleElementData(sb); break;
      case 0x13b: re.description = 'notify me when it is delivered'; re.data = new SimpleRuleElementData(sb); break;
      case 0x13c: re.description = 'Cc the message to people'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x13e: re.description = 'defer delivery by minutes'; re.data = new DeferDeliveryRuleElementData(sb); break;
      case 0x13f: re.description = 'perform custom action'; re.data = new PerformCustomActionRuleElementData(sb); break;
      case 0x142: re.description = 'stop processing more rules'; re.data = new SimpleRuleElementData(sb); break;
      case 0x144: re.description = 'redirect it to people'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x146: re.description = 'have server reply using message'; re.data = new AutomaticReplyRuleElementData(sb); break;
      case 0x147: re.description = 'forward as attachment'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x148: re.description = 'print it'; re.data = new SimpleRuleElementData(sb); break;
      case 0x149: re.description = 'start application'; re.data = new PathRuleElementData(sb); break;
      case 0x14a: re.description = 'permanently delete it'; re.data = new SimpleRuleElementData(sb); break;
      case 0x14b: re.description = 'run script'; re.data = new RunScriptRuleElementData(sb); break;
      case 0x14c: re.description = 'mark as read'; re.data = new SimpleRuleElementData(sb); break;
      case 0x14f: re.description = 'display a Desktop alert'; re.data = new SimpleRuleElementData(sb); break;
      case 0x151: re.description = 'flag for follow up'; re.data = new FlagForFollowUpRuleElementData(sb); break;
      case 0x152: re.description = "clear message's categories"; re.data = new SimpleRuleElementData(sb); break;
      case 0x153: re.description = 'apply retention policy'; re.data = new ApplyRetentionPolicyRuleElementData(sb); break;
      // === EXCEPTIONS (0x1F4-0x21B) ===
      case 0x1f4: re.description = 'except where my name is in the To box'; re.data = new SimpleRuleElementData(sb); break;
      case 0x1f5: re.description = 'except if sent only to me'; re.data = new SimpleRuleElementData(sb); break;
      case 0x1f6: re.description = 'except where my name is not in the To box'; re.data = new SimpleRuleElementData(sb); break;
      case 0x1f7: re.description = 'except if from <people or public group>'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x1f8: re.description = 'except if sent to <people or public group>'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x1f9: re.description = 'except with specific words in the subject'; re.data = new StringsListRuleElementData(sb); break;
      case 0x1fa: re.description = 'except with specific words in the body'; re.data = new StringsListRuleElementData(sb); break;
      case 0x1fb: re.description = 'except with specific words in the subject or body'; re.data = new StringsListRuleElementData(sb); break;
      case 0x1fc: re.description = 'except if flagged for <action>'; re.data = new FlaggedForActionRuleElementData(sb); break;
      case 0x1fe: re.description = 'except if marked as importance'; re.data = new ImportanceRuleElementData(sb); break;
      case 0x1ff: re.description = 'except if marked as sensitivity'; re.data = new SensitivityRuleElementData(sb); break;
      case 0x203: re.description = 'except if assigned to <category> category'; re.data = new CategoriesListRuleElementData(sb); break;
      case 0x208: re.description = 'except if it is an automatic reply'; re.data = new SimpleRuleElementData(sb); break;
      case 0x20a: re.description = 'except if it has an attachment'; re.data = new SimpleRuleElementData(sb); break;
      case 0x20b: re.description = 'except with selected properties of documents or forms'; re.data = new WithSelectedPropertiesOfDocumentOrFormsRuleElementData(sb); break;
      case 0x20c: re.description = 'except with a size in a specific range'; re.data = new SizeInSpecificRangeRuleElementData(sb); break;
      case 0x20d: re.description = 'except received in a specific date span'; re.data = new ReceivedInSpecificDateSpanRuleElementData(sb); break;
      case 0x20e: re.description = 'except where my name is in the CC box'; re.data = new SimpleRuleElementData(sb); break;
      case 0x20f: re.description = 'except where my name is in the To or CC box'; re.data = new SimpleRuleElementData(sb); break;
      case 0x210: re.description = 'except if it uses the <form> form'; re.data = new FormTypeRuleElementData(sb); break;
      case 0x211: re.description = 'except with specific words in the sender address'; re.data = new StringsListRuleElementData(sb); break;
      case 0x212: re.description = 'except with specific words in the recipient address'; re.data = new StringsListRuleElementData(sb); break;
      case 0x213: re.description = 'except with specific words in the message header'; re.data = new StringsListRuleElementData(sb); break;
      case 0x214: re.description = 'except through the specified account'; re.data = new ThroughAccountRuleElementData(sb); break;
      case 0x215: re.description = 'except if sender is in specified address book'; re.data = new SenderInSpecifiedAddressBookRuleElementData(sb); break;
      case 0x216: re.description = 'except on this machine only'; re.data = new SimpleRuleElementData(sb); break;
      case 0x218: re.description = 'except if it uses the <form> form (alt)'; re.data = new FormTypeRuleElementData(sb); break;
      case 0x219: re.description = 'except with specific words in the message header (alt)'; re.data = new StringsListRuleElementData(sb); break;
      case 0x21a: re.description = 'except if assigned to any category'; re.data = new SimpleRuleElementData(sb); break;
      case 0x21b: re.description = 'except if it is a meeting invitation or update'; re.data = new SimpleRuleElementData(sb); break;
      default: {
        re.description = `unknown element (0x${re.id.toString(16)})`;
        console.warn(`Warning: unknown rule element id 0x${re.id.toString(16)} at offset ${sb.offset}`);
        try { re.data = new SimpleRuleElementData(sb); } catch (e) { re.description += ' (could not parse data)'; }
        break;
      }
    }
    return re;
  }
}

class Rule {
  private header?: RuleHeader;
  private elements: RuleElement[] = [];
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

class RulesFile {
  private header?: RulesHeader;
  private rules: Rule[] = [];
  private footer?: RulesFooter;
  private constructor() {}
  public static parse(buf: StreamBuffer) {
    const rf = new RulesFile();
    rf.header = RulesHeader.parse(buf);
    console.log(`Parsing ${rf.header.numberOfRules} rules...`);
    for (let i = 0; i < rf.header.numberOfRules; i++) {
      try {
        const rule = Rule.parse(buf, i, rf.header.numberOfRules);
        rf.rules.push(rule);
      } catch (e) {
        console.warn(`Warning: could not parse rule ${i}: ${e}`);
        fs.writeFileSync('outlook-rules.json', JSON.stringify(rf, null, 2));
        console.log(`Wrote partial results (${rf.rules.length} rules) to outlook-rules.json`);
        break;
      }
      if (i !== rf.header.numberOfRules - 1) {
        try {
          const separator = buf.readUInt16();
          softAssert(separator === 0, `expected inter-rule separator 0, got 0x${separator.toString(16)}`);
        } catch (e) { break; }
      }
    }
    try { rf.footer = RulesFooter.parse(buf); } catch (e) { console.warn(`Warning: could not parse footer: ${e}`); }
    return rf;
  }
}

const content = fs.readFileSync(process.argv[2] || 'Untitled.rwz');

async function main() {
  const rf = await RulesFile.parse(new StreamBuffer(content));
  const outFile = 'outlook-rules.json';
  fs.writeFileSync(outFile, JSON.stringify(rf, null, 2));
  console.log(`Wrote ${outFile}`);
}

try {
  await main();
  process.exit(0);
} catch (e) {
  console.log(`Error ${e}`);
  process.exit(1);
}
