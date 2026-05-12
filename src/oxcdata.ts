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
        case OXCDATA.PtypGuid: {
          const offset = ph.data[1];
          softAssert(offset >= 0); softAssert(offset <= endPosition);
          sb.offset = startPos + offset;
          value = sb.readAsciiString(16);
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
      this.properties[ph.id] = value;
    }
    sb.offset = endPosition;
  }
}
