import { StreamBuffer } from './stream-buffer.js';
import { OutlookRulesReadError } from './errors.js';
import {
  RuleElementData, UnknownRuleElement0x64Data, ApplyRuleElementData,
  SimpleRuleElementData, PeopleOrPublicGroupListRuleElementData,
  StringsListRuleElementData, FlaggedForActionRuleElementData,
  MoveToFolderRuleElementData, ImportanceRuleElementData,
  SensitivityRuleElementData, CategoriesListRuleElementData,
  PathRuleElementData, DisplayMessageInNewItemAlertWindowRuleElementData,
  FlagRuleElementData, DeferDeliveryRuleElementData,
  PerformCustomActionRuleElementData, AutomaticReplyRuleElementData,
  RunScriptRuleElementData, FlagForFollowUpRuleElementData,
  ApplyRetentionPolicyRuleElementData, OnThisComputerOnlyRuleElementData,
  WithSelectedPropertiesOfDocumentOrFormsRuleElementData,
  SizeInSpecificRangeRuleElementData, ReceivedInSpecificDateSpanRuleElementData,
  FormTypeRuleElementData, ThroughAccountRuleElementData,
  SenderInSpecifiedAddressBookRuleElementData,
} from './rule-element-data.js';

export class RuleElement {
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
      case 0xcd: re.description = 'with <specific words> in the subject'; re.data = new StringsListRuleElementData(sb); break;
      case 0xce: re.description = 'with <specific words> in the body'; re.data = new StringsListRuleElementData(sb); break;
      case 0xcf: re.description = 'with <specific words> in the subject or body'; re.data = new StringsListRuleElementData(sb); break;
      case 0xd0: re.description = 'flagged for <action>'; re.data = new FlaggedForActionRuleElementData(sb); break;
      case 0xd2: re.description = 'marked as <importance>'; re.data = new ImportanceRuleElementData(sb); break;
      case 0xd3: re.description = 'marked as <sensitivity>'; re.data = new SensitivityRuleElementData(sb); break;
      case 0xd7: re.description = 'assigned to <category> category'; re.data = new CategoriesListRuleElementData(sb); break;
      case 0xdc: re.description = 'which is an automatic reply'; re.data = new SimpleRuleElementData(sb); break;
      case 0xde: re.description = 'which has an attachment'; re.data = new SimpleRuleElementData(sb); break;
      case 0xdf: re.description = 'with <selected properties> of documents or forms'; re.data = new WithSelectedPropertiesOfDocumentOrFormsRuleElementData(sb); break;
      case 0xe0: re.description = 'with a size <in a specific range>'; re.data = new SizeInSpecificRangeRuleElementData(sb); break;
      case 0xe1: re.description = 'received <in a specific date span>'; re.data = new ReceivedInSpecificDateSpanRuleElementData(sb); break;
      case 0xe2: re.description = 'where my name is in the Cc box'; re.data = new SimpleRuleElementData(sb); break;
      case 0xe3: re.description = 'where my name is in the To or Cc box'; re.data = new SimpleRuleElementData(sb); break;
      case 0xe4: re.description = 'uses the <form name> form'; re.data = new FormTypeRuleElementData(sb); break;
      case 0xe5: re.description = "with <specific words> in the recipient's address"; re.data = new StringsListRuleElementData(sb); break;
      case 0xe6: re.description = "with <specific words> in the sender's address"; re.data = new StringsListRuleElementData(sb); break;
      case 0xe8: re.description = 'with <specific words> in the message header'; re.data = new StringsListRuleElementData(sb); break;
      case 0xe9: re.description = 'from senders on my <Exception List>'; re.data = new SimpleRuleElementData(sb); break;
      case 0xeb: re.description = 'suspected to be junk email or from <Junk Senders>'; re.data = new SimpleRuleElementData(sb); break;
      case 0xec: re.description = 'containing adult content or from <Adult Content Senders>'; re.data = new SimpleRuleElementData(sb); break;
      case 0xed: re.description = 'with a relevance <in a specific range>'; re.data = new SizeInSpecificRangeRuleElementData(sb); break;
      case 0xee: re.description = 'through the <specified> account'; re.data = new ThroughAccountRuleElementData(sb); break;
      case 0xef: re.description = 'on this computer only'; re.data = new OnThisComputerOnlyRuleElementData(sb); break;
      case 0xf0: re.description = 'sender is in <specified> Address Book'; re.data = new SenderInSpecifiedAddressBookRuleElementData(sb); break;
      case 0xf1: re.description = 'which is a meeting invitation or update'; re.data = new SimpleRuleElementData(sb); break;
      case 0xf2: re.description = 'from my contacts'; re.data = new SimpleRuleElementData(sb); break;
      case 0xf3: re.description = 'from a subscription'; re.data = new SimpleRuleElementData(sb); break;
      case 0xf4: re.description = 'of the <specific> form type'; re.data = new FormTypeRuleElementData(sb); break;
      case 0xf5: re.description = 'from RSS feeds with <specified text> in the title'; re.data = new StringsListRuleElementData(sb); break;
      case 0xf6: re.description = 'assigned to any category'; re.data = new SimpleRuleElementData(sb); break;
      case 0xf7: re.description = 'from any RSS feed'; re.data = new SimpleRuleElementData(sb); break;

      // === ACTIONS (0x12C-0x153) ===
      case 0x12c: re.description = 'move it to the <specified> folder'; re.data = new MoveToFolderRuleElementData(sb); break;
      case 0x12d: re.description = 'delete it'; re.data = new SimpleRuleElementData(sb); break;
      case 0x12e: re.description = 'forward it to <people or public group>'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x12f: re.description = 'reply using <template>'; re.data = new PathRuleElementData(sb); break;
      case 0x130: re.description = 'display <a specific message> in the New Item Alert window'; re.data = new DisplayMessageInNewItemAlertWindowRuleElementData(sb); break;
      case 0x131: re.description = 'flag message for <action in a number of days>'; re.data = new FlagRuleElementData(sb); break;
      case 0x132: re.description = 'clear the Message flag'; re.data = new SimpleRuleElementData(sb); break;
      case 0x133: re.description = 'assign it to the <category> category'; re.data = new CategoriesListRuleElementData(sb); break;
      case 0x136: re.description = 'play <sound>'; re.data = new PathRuleElementData(sb); break;
      case 0x137: re.description = 'mark it as <importance>'; re.data = new ImportanceRuleElementData(sb); break;
      case 0x138: re.description = 'mark it as <sensitivity>'; re.data = new SensitivityRuleElementData(sb); break;
      case 0x139: re.description = 'move a copy to the <specified> folder'; re.data = new MoveToFolderRuleElementData(sb); break;
      case 0x13a: re.description = 'notify me when it is read'; re.data = new SimpleRuleElementData(sb); break;
      case 0x13b: re.description = 'notify me when it is delivered'; re.data = new SimpleRuleElementData(sb); break;
      case 0x13c: re.description = 'Cc the message to <people or public group>'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x13e: re.description = 'defer delivery by <a number of> minutes'; re.data = new DeferDeliveryRuleElementData(sb); break;
      case 0x13f: re.description = 'perform <a custom action>'; re.data = new PerformCustomActionRuleElementData(sb); break;
      case 0x142: re.description = 'stop processing more rules'; re.data = new SimpleRuleElementData(sb); break;
      case 0x143: re.description = 'do not search message for commercial or adult content'; re.data = new SimpleRuleElementData(sb); break;
      case 0x144: re.description = 'redirect it to <people or public group>'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x145: re.description = 'add <number> to relevance'; re.data = new SimpleRuleElementData(sb); break;
      case 0x146: re.description = 'have server reply using <a specific message>'; re.data = new AutomaticReplyRuleElementData(sb); break;
      case 0x147: re.description = 'forward it to <people or public group> as attachment'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x148: re.description = 'print it'; re.data = new SimpleRuleElementData(sb); break;
      case 0x149: re.description = 'start <application>'; re.data = new PathRuleElementData(sb); break;
      case 0x14a: re.description = 'permanently delete it'; re.data = new SimpleRuleElementData(sb); break;
      case 0x14b: re.description = 'run <script>'; re.data = new RunScriptRuleElementData(sb); break;
      case 0x14c: re.description = 'mark as read'; re.data = new SimpleRuleElementData(sb); break;
      case 0x14f: re.description = 'display a Desktop alert'; re.data = new SimpleRuleElementData(sb); break;
      case 0x150: re.description = 'color flag'; re.data = new SimpleRuleElementData(sb); break;
      case 0x151: re.description = 'flag for follow up'; re.data = new FlagForFollowUpRuleElementData(sb); break;
      case 0x152: re.description = "clear message's categories"; re.data = new SimpleRuleElementData(sb); break;
      case 0x153: re.description = 'apply retention policy'; re.data = new ApplyRetentionPolicyRuleElementData(sb); break;

      // === EXCEPTIONS (0x1F4-0x21B) ===
      case 0x1f4: re.description = 'except where my name is in the To box'; re.data = new SimpleRuleElementData(sb); break;
      case 0x1f5: re.description = 'except if sent only to me'; re.data = new SimpleRuleElementData(sb); break;
      case 0x1f6: re.description = 'except where my name is not in the To box'; re.data = new SimpleRuleElementData(sb); break;
      case 0x1f7: re.description = 'except if from <people or public group>'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x1f8: re.description = 'except if sent to <people or public group>'; re.data = new PeopleOrPublicGroupListRuleElementData(sb); break;
      case 0x1f9: re.description = 'except if the subject contains <specific words>'; re.data = new StringsListRuleElementData(sb); break;
      case 0x1fa: re.description = 'except if the body contains <specific words>'; re.data = new StringsListRuleElementData(sb); break;
      case 0x1fb: re.description = 'except if the subject or body contains <specific words>'; re.data = new StringsListRuleElementData(sb); break;
      case 0x1fc: re.description = 'except if it is flagged for <action>'; re.data = new FlaggedForActionRuleElementData(sb); break;
      case 0x1fe: re.description = 'except if it is marked as <importance>'; re.data = new ImportanceRuleElementData(sb); break;
      case 0x1ff: re.description = 'except if it is marked as <sensitivity>'; re.data = new SensitivityRuleElementData(sb); break;
      case 0x203: re.description = 'except if it is assigned to <category> category'; re.data = new CategoriesListRuleElementData(sb); break;
      case 0x208: re.description = 'except if it is an automatic reply'; re.data = new SimpleRuleElementData(sb); break;
      case 0x20a: re.description = 'except if it has an attachment'; re.data = new SimpleRuleElementData(sb); break;
      case 0x20b: re.description = 'except with <selected properties> of documents or forms'; re.data = new WithSelectedPropertiesOfDocumentOrFormsRuleElementData(sb); break;
      case 0x20c: re.description = 'except with a size <in a specific range>'; re.data = new SizeInSpecificRangeRuleElementData(sb); break;
      case 0x20d: re.description = 'except if received <in a specific date span>'; re.data = new ReceivedInSpecificDateSpanRuleElementData(sb); break;
      case 0x20e: re.description = 'except where my name is in the Cc box'; re.data = new SimpleRuleElementData(sb); break;
      case 0x20f: re.description = 'except if my name is in the To or Cc box'; re.data = new SimpleRuleElementData(sb); break;
      case 0x210: re.description = 'except if it uses the <form name> form'; re.data = new FormTypeRuleElementData(sb); break;
      case 0x211: re.description = "except with <specific words> in the recipient's address"; re.data = new StringsListRuleElementData(sb); break;
      case 0x212: re.description = "except with <specific words> in the sender's address"; re.data = new StringsListRuleElementData(sb); break;
      case 0x213: re.description = 'except if the message header contains <specific words>'; re.data = new StringsListRuleElementData(sb); break;
      case 0x214: re.description = 'except through the <specified> account'; re.data = new ThroughAccountRuleElementData(sb); break;
      case 0x215: re.description = 'except if sender is in <specified> Address Book'; re.data = new SenderInSpecifiedAddressBookRuleElementData(sb); break;
      case 0x216: re.description = 'except if it is a meeting invitation or update'; re.data = new SimpleRuleElementData(sb); break;
      case 0x217: re.description = 'except from my contacts'; re.data = new SimpleRuleElementData(sb); break;
      case 0x218: re.description = 'except if it is of the <specific> form type'; re.data = new FormTypeRuleElementData(sb); break;
      case 0x219: re.description = 'except from RSS feeds with <specified text> in the title'; re.data = new StringsListRuleElementData(sb); break;
      case 0x21a: re.description = 'except if assigned to any category'; re.data = new SimpleRuleElementData(sb); break;
      case 0x21b: re.description = 'except from any RSS feed'; re.data = new SimpleRuleElementData(sb); break;

      default: {
        throw new OutlookRulesReadError(`unknown element data type: 0x${re.id!.toString(16)} (${re.id})`);
      }
    }
    return re;
  }
}
