import { StreamBuffer } from './stream-buffer.js';
import { softAssert } from './errors.js';

export enum OXCDATA {
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

/** Well-known MAPI property IDs from MS-OXPROPS */
const MAPI_PROPERTY_NAMES: Record<number, string> = {
  0x0001: 'TemplateData',
  0x0002: 'AlternateRecipientAllowed',
  0x000f: 'DeferredDeliveryTime',
  0x0017: 'Importance',
  0x001a: 'MessageClass',
  0x0036: 'Sensitivity',
  0x0037: 'Subject',
  0x0039: 'ClientSubmitTime',
  0x003b: 'SentRepresentingSearchKey',
  0x003d: 'SubjectPrefix',
  0x0042: 'SentRepresentingName',
  0x0c15: 'RecipientType',
  0x0c24: 'RecipientEntryId', // undocumented but common
  0x0c29: 'ReplyRecipientSmtpProxies', // undocumented
  0x0c2c: 'RecipientOrder', // undocumented
  0x0e01: 'DeleteAfterSubmit',
  0x0e02: 'DisplayBcc',
  0x0e03: 'DisplayCc',
  0x0e04: 'DisplayTo',
  0x0e0f: 'RecipientReassignmentProhibited',
  0x0ff6: 'InstanceKey',
  0x0ff9: 'RecordKey',
  0x0ffe: 'ObjectType',
  0x0fff: 'EntryId',
  0x3000: 'RowId',
  0x3001: 'DisplayName',
  0x3002: 'AddressType',
  0x3003: 'EmailAddress',
  0x300b: 'SearchKey',
  0x3900: 'DisplayType',
  0x3905: 'DisplayTypeEx',
  0x39fe: 'SmtpAddress',
  0x39ff: 'AddressBookPhoneticDisplayName',
  0x3a00: 'Account',
  0x3a20: 'TransmittableDisplayName',
  0x3a40: 'SendRichInfo',
  0x3d01: 'MailPermission',
  0x5fde: 'RecipientResourceState',
  0x5fdf: 'RecipientOrder2',
  0x5fe5: 'RecipientSipUri',
  0x5feb: 'RecipientDisplayName2',
  0x5fef: 'RecipientSmtpAddress',
  0x5ff2: 'RecipientAddressType',
  0x5ff5: 'RecipientEntryId2',
  0x5ff6: 'RecipientFlags',
  0x5ff7: 'RecipientTrackStatus',
  0x5ffd: 'RecipientProposed',
  0x5fff: 'RecipientDisplayType',
};

/** Get a human-readable property key name */
export function getPropertyName(id: number): string {
  return MAPI_PROPERTY_NAMES[id] ?? `0x${id.toString(16).padStart(4, '0')}`;
}

/** Format a GUID from a 16-byte buffer into standard form */
function formatGuid(raw: Buffer): string {
  const hex = (offset: number, len: number) =>
    raw.subarray(offset, offset + len).toString('hex');
  // GUIDs are stored as: Data1(4 LE) - Data2(2 LE) - Data3(2 LE) - Data4(8)
  const d1 = raw.readUInt32LE(0).toString(16).padStart(8, '0');
  const d2 = raw.readUInt16LE(4).toString(16).padStart(4, '0');
  const d3 = raw.readUInt16LE(6).toString(16).padStart(4, '0');
  const d4 = hex(8, 2);
  const d5 = hex(10, 6);
  return `${d1}-${d2}-${d3}-${d4}-${d5}`;
}

export class PropertyValueHeader {
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

export class PropertyValueArray {
  public propertyHeaders: PropertyValueHeader[] = [];
  public readonly properties: Record<string, any> = {};

  public toJSON() { return this.properties; }

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
          value = sb.readBytes(length).toString('hex');
          sb.offset = position;
          break;
        }
        case OXCDATA.PtypBoolean: { value = ph.data[1] !== 0; break; }
        case OXCDATA.PtypGuid: {
          const offset = ph.data[1];
          softAssert(offset >= 0); softAssert(offset <= endPosition);
          sb.offset = startPos + offset;
          const guidBytes = sb.readBytes(16);
          value = formatGuid(guidBytes);
          sb.offset = position;
          break;
        }
        case OXCDATA.PtypTime: {
          const offset = ph.data[1];
          sb.offset = startPos + offset;
          value = sb.readDouble();
          sb.offset = position;
          break;
        }
        default: { throw new Error(`Property data type not yet implemented: ${OXCDATA[ph.dataType] ?? ph.dataType} (0x${ph.dataType.toString(16)})`); }
      }
      const key = getPropertyName(ph.id);
      this.properties[key] = value;
    }
    sb.offset = endPosition;
  }
}
