import { StreamBuffer } from './stream-buffer.js';
import { OutlookRulesReadError, softAssert } from './errors.js';
import { PropertyValueArray } from './oxcdata.js';

// Base class — toJSON strips protocol boilerplate (extended, reserved)
const HIDDEN_FIELDS = new Set(['extended', 'reserved']);

export class RuleElementData {
  public constructor(_sb: StreamBuffer) {}

  public toJSON(): Record<string, any> {
    const result: Record<string, any> = {};
    for (const [key, value] of Object.entries(this)) {
      if (!HIDDEN_FIELDS.has(key)) {
        result[key] = value;
      }
    }
    return result;
  }
}

// --- Mandatory elements ---

export class UnknownRuleElement0x64Data extends RuleElementData {
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
  public toJSON(): Record<string, any> { return { flags: this.flags }; }
}

export class ApplyRuleElementData extends UnknownRuleElement0x64Data {
  public when() {
    switch (this.flags) {
      case 0x1: return 'after the message arrives';
      case 0x4: return 'after I send the message';
      case 0x8: return 'after the server receives the message';
      default: return `unknown (0x${this.flags.toString(16)})`;
    }
  }
  public override toJSON(): Record<string, any> { return { appliesTo: this.when() }; }
}

export class SimpleRuleElementData extends RuleElementData {
  public extended?: number;
  public constructor(sb: StreamBuffer) {
    super(sb);
    this.extended = sb.readUInt32();
    softAssert(this.extended === 0);
  }
}

// --- Entry ID helpers ---

export class FlatEntry {
  public readonly size: number;
  public constructor(sb: StreamBuffer) { this.size = sb.readUInt32(); }
}

export class FolderEntryId extends FlatEntry {
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
      this.providerUID = sb.readBytes(16).toString('hex');
      this.folderType = sb.readUInt16();
      this.databaseGuid = sb.readBytes(16).toString('hex');
      this.globalCounter = sb.readUInt64().toString();
    }
    sb.offset = pos + this.size + 4;
  }
}

export class StoreEntryId extends FlatEntry {
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
    this.providerUID = sb.readBytes(16).toString('hex');
    this.version = sb.readUInt8();
    softAssert(this.version === 0);
    this.flag = sb.readUInt8();
    softAssert(this.flags === 0);
    this.dllFileName = sb.readAsciiString(14);
    softAssert(this.dllFileName === 'EMSMDB.DLL');
    this.wrappedFlags = sb.readUInt32();
    softAssert(this.wrappedFlags === 0);
    this.wrappedProvider = sb.readBytes(16).toString('hex');
    this.wrappedType = sb.readUInt32();
    softAssert(this.wrappedType === 0xc);
    this.serverShortName = sb.readAsciiUntilNullTerminator();
    this.mailboxDN = sb.readAsciiUntilNullTerminator();
    sb.offset = pos + this.size + 4;
  }
}

// --- Complex data types ---

export class PeopleOrPublicGroupListRuleElementData extends RuleElementData {
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
  public toJSON() { return { recipients: this.values }; }
}

export class SearchEntry {
  public flags?: number;
  public value?: string;
  public constructor(sb: StreamBuffer) {
    this.flags = sb.readUInt32();
    softAssert(this.flags === 0);
    this.value = sb.readStringObject();
  }
  public toJSON() { return this.value; }
}

export class StringsListRuleElementData extends RuleElementData {
  public entries: SearchEntry[] = [];
  public constructor(sb: StreamBuffer) {
    super(sb);
    const nEntries = sb.readUInt32();
    for (let i = 0; i < nEntries; i++) { this.entries.push(new SearchEntry(sb)); }
  }
  public toJSON() { return { words: this.entries }; }
}

export class FlaggedForActionRuleElementData extends RuleElementData {
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

export class MoveToFolderRuleElementData extends RuleElementData {
  public readonly extended: number;
  public readonly reserved: number;
  public readonly folderEntryId: FolderEntryId;
  public readonly storeEntryId: StoreEntryId;
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
  public toJSON() {
    return {
      folderName: this.folderName,
      mailbox: this.storeEntryId.serverShortName,
      secondaryUserStore: this.secondaryUserStore,
    };
  }
}

// --- Extended/Reserved + payload data types ---

export class ImportanceRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public importance: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.importance = sb.readUInt32();
  }
  public toJSON() { return { importance: (['Low', 'Normal', 'High'] as const)[this.importance] ?? this.importance }; }
}

export class SensitivityRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public sensitivity: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.sensitivity = sb.readUInt32();
  }
  public toJSON() { return { sensitivity: (['Normal', 'Personal', 'Private', 'Confidential'] as const)[this.sensitivity] ?? this.sensitivity }; }
}

export class CategoriesListRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public categories: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.categories = sb.readStringObject();
  }
}

export class PathRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public path: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.path = sb.readStringObject();
  }
}

export class DisplayMessageInNewItemAlertWindowRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public message: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.message = sb.readStringObject();
  }
}

export class FlagRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public days: number;
  public actionName: string; public unknown: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.days = sb.readUInt32(); this.actionName = sb.readStringObject();
    this.unknown = sb.readUInt32();
  }
}

export class DeferDeliveryRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public minutes: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.minutes = sb.readUInt32();
  }
}

export class PerformCustomActionRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public location: string; public name: string; public options: string; public actionValue: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.location = sb.readStringObject(); this.name = sb.readStringObject();
    this.options = sb.readStringObject(); this.actionValue = sb.readStringObject();
  }
}

export class AutomaticReplyRuleElementData extends RuleElementData {
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

export class RunScriptRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public name: string; public functionName: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.name = sb.readStringObject(); this.functionName = sb.readStringObject();
  }
}

export class FlagForFollowUpRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public followUp: number; public actionName: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.followUp = sb.readUInt32(); this.actionName = sb.readStringObject();
  }
}

export class ApplyRetentionPolicyRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public followUp: number; public guid: string; public name: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.followUp = sb.readUInt32(); this.guid = sb.readBytes(16).toString('hex');
    this.name = sb.readStringObject();
  }
}

export class OnThisComputerOnlyRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public uuid: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.uuid = sb.readBytes(16).toString('hex');
  }
}

export class WithSelectedPropertiesOfDocumentOrFormsRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
  }
}

export class SizeInSpecificRangeRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public minSize: number; public maxSize: number;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.minSize = sb.readUInt32(); this.maxSize = sb.readUInt32();
  }
}

export class ReceivedInSpecificDateSpanRuleElementData extends RuleElementData {
  public extended: number; public reserved: number;
  public startDate: any; public endDate: any;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    const startStatus = sb.readUInt32();
    this.startDate = { status: startStatus === 0 ? 'Valid' : 'Null', timestamp: sb.readDouble() };
    const endStatus = sb.readUInt32();
    this.endDate = { status: endStatus === 0 ? 'Valid' : 'Null', timestamp: sb.readDouble() };
  }
}

export class FormTypeRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public formClass: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.formClass = sb.readStringObject();
  }
}

export class ThroughAccountRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public accountName: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.accountName = sb.readStringObject();
  }
}

export class SenderInSpecifiedAddressBookRuleElementData extends RuleElementData {
  public extended: number; public reserved: number; public addressBookName: string;
  public constructor(sb: StreamBuffer) {
    super(sb); this.extended = sb.readUInt32(); softAssert(this.extended === 1);
    this.reserved = sb.readUInt32(); softAssert(this.reserved === 0);
    this.addressBookName = sb.readStringObject();
  }
}
