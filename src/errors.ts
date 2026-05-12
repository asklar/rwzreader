export class OutlookRulesReadError extends Error {}

export function softAssert(condition: boolean, msg?: string) {
  if (!condition) {
    console.warn(`Warning: assertion failed${msg ? ': ' + msg : ''}`);
  }
}
