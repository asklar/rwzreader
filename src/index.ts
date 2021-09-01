import assert, { rejects } from 'assert';
import * as fs from 'fs';
import { Readable } from 'stream';
import { promisify } from 'util';

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
    const startPos = this.offset;
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
    const origPos = this.offset;
    const v = Buffer.alloc(260);
    try {
      let i = 0;
      for (let c = this.readUInt16(); c !== 0; c = this.readUInt16()) {
        v.writeUInt16LE(c, i++)
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
      case 1310720:
        rh.version = 'outlook2019'; break;
      case 1200000:
        rh.version = 'outlook2007'; break;
      case 1100000:
        rh.version = 'outlook2003'; break;
      case 1000000:
        rh.version = 'outlook2002'; break;
      case 980413:
        rh.version = 'outlook2000'; break;
      case 970812:
        rh.version = 'outlook98'; break;
      case 0:
        rh.version = 'noSignatureOutlook2003'; break;
      default:
        rh.version = 'noSignature'; break;
    }


    if (rh.version != 'noSignatureOutlook2003' && rh.version != 'noSignature') {
      /// Signature (4 bytes)
      const signature = peekedSignature;
      rh.signature = signature

      /// Flags (4 bytes)
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
          }
          else {
            throw new OutlookRulesReadError('invalidSignature'); //(signature: signature, flags: flags)
          }
        } catch (e) {
          console.log(e);
        }
      } else {
        rh.flags = undefined;
      }
    } else {
      rh.signature = undefined;
      rh.flags = undefined;
    }

    /// Unknown1 (4 bytes)
    if (rh.version !== 'noSignature') {
      rh.unknown[1] = buf.readUInt32();
      if (rh.unknown[1] === 0x00000000) { } else {
        throw new OutlookRulesReadError('corrupted 1');
      }
    } else {
      rh.unknown[1] = undefined;
    }

    /// Unknown2 (4 bytes)
    if (rh.version !== 'noSignature') {
      rh.unknown[2] = buf.readUInt32();
      if (rh.unknown[2] === 0x00000000) { } else {
        throw new OutlookRulesReadError('corrupted 2');
      }
    } else {
      rh.unknown[2] = undefined;
    }

    /// unknown[3] (4 bytes)
    if (rh.version != 'noSignature') {
      rh.unknown[3] = buf.readUInt32();
      if (rh.unknown[3] == 0x00000000) { } else {
        throw new OutlookRulesReadError('corrupted 3');
      }
    } else {
      rh.unknown[3] = undefined;
    }

    /// Unknown4 (4 bytes)
    if (rh.version != 'noSignature') {
      rh.unknown[4] = buf.readUInt32();
    } else {
      rh.unknown[4] = undefined;
    }

    /// Unknown5 (4 bytes)
    if (rh.version != 'noSignature') {
      rh.unknown[5] = buf.readUInt32();
    } else {
      rh.unknown[6] = undefined;
    }

    /// Unknown6 (4 bytes)
    if (rh.version != 'noSignature') {
      rh.unknown[6] = buf.readUInt32();
    } else {
      rh.unknown[6] = undefined;
    }

    /// Unknown7 (4 bytes)
    if (rh.version != 'noSignature') {
      rh.unknown[7] = buf.readUInt32();
    } else {
      rh.unknown[7] = undefined;
    }

    /// Unknown8 (4 bytes)
    if (rh.version != 'noSignature') {
      rh.unknown[8] = buf.readUInt32();
      if (rh.unknown[8] == 0x00000001) { } else {
        throw new OutlookRulesReadError('corrupted 8');
      }
    } else {
      rh.unknown[9] = undefined;
    }

    /// Unknown9 (4 bytes)
    if (rh.version >= 'outlook2002' || rh.version == 'noSignatureOutlook2003') {
      rh.unknown[9] = buf.readUInt32();
    } else {
      rh.unknown[9] = undefined;
    }

    /// Number of Rules (2 bytes)
    rh.numberOfRules = buf.readUInt16();

    const extra = buf.readUInt16();
    return rh;
  }

  private constructor() { }
}

class OleDateTime {
  private status?: 'Valid' | 'Null';
  private timestamp?: number;
  
  public get createdOn() : Date {
    return new Date(1900, 1, this.timestamp);
  }
  
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

  private data?: string;

  private constructor() {}
  public static parse(sb: StreamBuffer, index: number) {
    const rh = new RuleHeader();
    rh.signature = sb.readUInt16();
    rh.name = sb.readStringObject();
    rh.enabled = sb.readUInt32() === 0 ? false : true;
    for (let i = 0; i < 4; i++) {
      rh.unknown[i] = sb.readUInt32();
    }
    rh.dataSize = sb.readUInt32();
    rh.nRuleElements = sb.readUInt16();

    rh.separator = sb.readUInt16();

    if (rh.separator === 0xffff) {
      assert(index === 0);
      rh.padding = sb.readUInt16();

      rh.classNameLength = sb.readUInt16();
      assert(rh.classNameLength === 'CRuleElement'.length);
      rh.className = sb.readAsciiString(rh.classNameLength);
      assert(rh.className === 'CRuleElement');  
    } else if (rh.separator === 0x8001) {
      assert(index !== 0);
    } else if (rh.separator === 0) {
      // ?????? 
    } else {
      throw new OutlookRulesReadError('corrupted separator');
    }


    // const remainingData = rh.dataSize - (2 + 2 + 2 + 2 + 'CRuleElement'.length);
    // rh.data = sb.readString(remainingData);
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
    assert(this.extended === 0x1);
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
    assert(this.extended === 0);
  }
}

enum OXCDATA {
  // this describes the dataType member in PropertyValueHeader
  // See section 2.11.1 of MS-OXCDATA:
  // https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/0c77892e-288e-435a-9c49-be1c20c7afdb

  PtypInteger16 = 2,
  PtypInteger32 = 3,
  PtypFloating32 = 4,
  PtypFloating64 = 5,
  PtypCurrency = 6,
  PtypFloatingTime = 7,
  PtypErrorCode = 0xa,
  PtypBoolean = 0xb,
  PtypInteger64 = 0x14,
  PtypString = 0x1f,
  PtypString8 = 0x1e,
  PtypTime = 0x40,
  PtypGuid = 0x48,
  PtypServerId = 0xfb,
  PtypRestriction = 0xfd,
  PtypRuleAction = 0xfe,
  PtypBinary = 0x102,
  PtypMultipleInteger16 = 0x1002,
  PtypMultipleInteger32 = 0x1003,
  PtypMultipleFloating32 = 0x1004,
  PtypMultipleFloating64 = 0x1005,
  PtypMultipleCurrency = 0x1006,
  PtypMultipleFloatingTime = 0x1007,
  PtypMultipleInteger64 = 0x1014,
  PtypMultipleString = 0x101f,
  PtypMultipleString8 = 0x101e,
  PtypMultipleTime = 0x1040,
  PtypMultipleGuid = 0x1048,
  PtypMultipleBinary = 0x1102,
  PtypUnspecified = 0,
  PtypNull = 1,
  PtypObject = 0xd,
  PtypEmbeddedTable = 0xd,
}

class PropertyValueHeader {
  public readonly dataType: OXCDATA;
  public readonly id: number;
  public readonly data: number[] = [];
  public type() {
    return OXCDATA[this.dataType];
  }
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
  public propertyData?: string;
  public readonly properties: Record<number, any> = {};
  public constructor(sb: StreamBuffer) {
    const unknown = sb.readUInt32();
    //assert(unknown === 0);
    const nProps = sb.readUInt32();
    const propDataSize = sb.readUInt32();

    const startPos = sb.offset;
    const endPosition = sb.offset + propDataSize;
    for (let i = 0; i < nProps; i++) {
      const ph = new PropertyValueHeader(sb);
      this.propertyHeaders.push(ph);

      const position = sb.offset;
      let value = undefined;
      switch (ph.dataType) {
        case OXCDATA.PtypInteger32:
          value = ph.data[1]; break;
        case OXCDATA.PtypErrorCode:
          value = ph.data[1]; break;
        case OXCDATA.PtypString: {
          const offset = ph.data[1];
          assert(offset >= 0);
          assert(offset <= endPosition);
          sb.offset = startPos + offset;
          value = sb.readStringUntilNullTerminator();
          assert(sb.offset <= endPosition);
          sb.offset = position;
          break;
        }
        case OXCDATA.PtypString8: {
          const offset = ph.data[1];
          assert(offset >= 0);
          assert(offset <= endPosition);
          sb.offset = startPos + offset;
          value = sb.readAsciiUntilNullTerminator();
          assert(sb.offset <= endPosition);
          sb.offset = position;
          break;
        }
        case OXCDATA.PtypBinary: {
          const offset = ph.data[1];
          assert(offset >= 0);
          assert(offset <= endPosition);
          sb.offset = startPos + offset;
          const length = ph.data[2];
          assert(length >= 0);
          assert(sb.offset + length <= endPosition);
          value = sb.readAsciiString(length);
          sb.offset = position;
          break;
        }
        case OXCDATA.PtypBoolean: {
          value = ph.data[1] !== 0;
          break;
        }
        default: {
          throw new Error('NYI');
        }
      }
      this.properties[ph.id] = value;
    }
    //this.propertyData = sb.readAsciiString(propDataSize - nProps * 12);
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
    assert(this.extended === 1);
    this.reserved = sb.readUInt32();
    assert(this.reserved === 0);
    const nValues = sb.readUInt32();
    for (let i = 0; i < nValues; i++) {
      const v = new PropertyValueArray(sb);
      this.values.push(v);
    }
    const unknown1 = sb.readUInt32();
    assert(unknown1 === 1);
    const unknown2 = sb.readUInt32();
    assert(unknown2 === 0);
  }
}

class SearchEntry {
  public flags?: number;
  public value?: string;
  public constructor(sb: StreamBuffer) {
    this.flags = sb.readUInt32();
    assert(this.flags === 0);
    this.value = sb.readStringObject();
  }
}
class StringsListRuleElementData extends RuleElementData {
  public entries: SearchEntry[] = [];
  public constructor(sb: StreamBuffer) {
    super(sb);
    const nEntries = sb.readUInt32();
    for (let i = 0; i < nEntries; i++) {
      const e = new SearchEntry(sb);
      this.entries.push(e);
    }
  }
}

class FlaggedForActionRuleElementData extends RuleElementData {

}

class FlatEntry {
  public readonly size: number;
  public constructor(sb: StreamBuffer) {
    this.size = sb.readUInt32();
  }
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
      assert(this.flags === 0);
      this.providerUID = sb.readAsciiString(16);
      this.folderType = sb.readUInt16();
      this.databaseGuid = sb.readAsciiString(16);
      this.globalCounter = sb.readUInt64().toString();
    } else {
      console.log('size == 0');
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
    assert(this.flags === 0);
    this.providerUID = sb.readAsciiString(16);
    this.version = sb.readUInt8();
    assert(this.version === 0);
    this.flag = sb.readUInt8();
    assert(this.flags === 0);
    this.dllFileName = sb.readAsciiString(14);
    assert(this.dllFileName === 'EMSMDB.DLL\0\0\0\0');
    this.wrappedFlags = sb.readUInt32();
    assert(this.wrappedFlags === 0);
    this.wrappedProvider = sb.readAsciiString(16);
    // const mailboxStoreObject = '\x1B\x55\xFA\x20\xAA\x66\x11\xCD\x9B\xC8\x00\xAA\x00\x2F\xC4\x5A';
    // assert(this.wrappedProvider === mailboxStoreObject);
    this.wrappedType = sb.readUInt32();
    assert(this.wrappedType === 0xc);
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
class RuleElementDescription {
  public constructor(public description: string, public data: RuleElementData) {}

}

// type RuleMap = { [id: number]: { description: string, factory: {abstract new <T extends RuleElementDescription>(sb:StreamBuffer): T} };

// const RuleElementId: RuleMap = {
//   0x64: { description: 'Unknown', factory: UnknownRuleElement0x64Data}

// };


class RuleElement {
  public id?: number;
  public description?: string;
  public data?: RuleElementData;
  private constructor(){}
  public static parse(sb: StreamBuffer) {
    const re = new RuleElement();
    re.id = sb.readUInt32();
    switch (re.id) {
      // Mandatory rule elements
      case 0x64: {
        re.description = 'Unknown';
        re.data = new UnknownRuleElement0x64Data(sb);
        break;
      }
      case 0x190: {
        re.description = 'type of message to which this rule applies';
        re.data = new ApplyRuleElementData(sb);
        break;
      }
      // Conditions
      case 0xc8: {
        re.description = 'where my name is in the To box';
        re.data = new SimpleRuleElementData(sb);
        break;
      }
      case 0xc9: {
        re.description = 'sent only to me';
        re.data = new SimpleRuleElementData(sb);
        break;
      } 
      case 0xca: {
        re.description = 'where my name is not in the To box';
        re.data = new SimpleRuleElementData(sb);
        break;
      }
      case 0xcb: {
        re.description = 'from <people or public group>';
        re.data = new PeopleOrPublicGroupListRuleElementData(sb);
        break;
      }
      case 0xcc: {
        re.description = 'sent to <people or public group>';
        re.data = new PeopleOrPublicGroupListRuleElementData(sb);
        break;
      }
      case 0xcd: {
        re.description = 'with specific words in the subject';
        re.data = new StringsListRuleElementData(sb);
        break;
      }
      case 0xce: {
        re.description = 'with specific words in the body';
        re.data = new StringsListRuleElementData(sb);
        break;
      }
      case 0xcf: {
        re.description = 'with specific words in the subject or body';
        re.data = new StringsListRuleElementData(sb);
        break;
      }
      case 0xd0: {
        re.description = 'flagged for <action>';
        re.data = new FlaggedForActionRuleElementData(sb);
        break;
      }

      case 0xe2: {
        re.description = 'where my name is in the CC box';
        re.data = new SimpleRuleElementData(sb);
        break;
      }

      case 0xe8: {
        re.description = 'with specific words in the message header';
        re.data = new StringsListRuleElementData(sb);
        break;
      }

      case 0xf6: {
        re.description = 'assigned to any category';
        re.data = new SimpleRuleElementData(sb);
        break;
      }

      case 0x12c: {
        re.description = 'move it to the specified folder';
        re.data = new MoveToFolderRuleElementData(sb);
        break;
      }
      case 0x142: {
        re.description = 'stop processing more rules';
        re.data = new SimpleRuleElementData(sb);
        break;
      }
      case 0x152: {
          re.description = 'clear message categories';
          re.data = new SimpleRuleElementData(sb);
          break;
      }

      default: {
        throw new OutlookRulesReadError(`unknown element data type: ${re.id}`);
      }
    }
    return re;
  }
}

class Rule {
  private header?: RuleHeader;
  private elements: RuleElement[] = [];
  private constructor() {}
  public static parse(sb: StreamBuffer, index: number) {
    const r = new Rule();
    r.header = RuleHeader.parse(sb, index);
    for (let i = 0; i < r.header.nRuleElements!; i++) {
      const elem = RuleElement.parse(sb);
      r.elements.push(elem);
      if (i !== r.header.nRuleElements! - 1) {
        const separator = sb.readUInt16();
        assert(separator === 0x8001);
      }
    }
    return r;
  }
}

class RulesFile {

  private header?: RulesHeader;
  private rules: Rule[] = [];
  private footer?: RulesFooter;
  private constructor() { }
  public static parse(buf: StreamBuffer) {
    const rf = new RulesFile();
    rf.header = RulesHeader.parse(buf);
    
    // console.log(JSON.stringify(rf.header, null, 2));

    for (let i = 0; i < rf.header.numberOfRules; i++) {
      const rule = Rule.parse(buf, i);
      rf.rules.push(rule);

      if (i !== rf.header.numberOfRules - 1) {
        const separator = buf.readUInt16();
        assert(separator === 0);
      }
    }

    rf.footer = RulesFooter.parse(buf);

    return rf;
  }

}

const content = fs.readFileSync('C:/Temp/Untitled.rwz');
async function main() {
  const rf = await RulesFile.parse(new StreamBuffer(content));
  fs.writeFileSync('C:/temp/outlook-rules.json', JSON.stringify(rf, null, 2));
}

try {
  await main();
  process.exit(0);
} catch (e) {
  console.log(`Error ${e}`);
}