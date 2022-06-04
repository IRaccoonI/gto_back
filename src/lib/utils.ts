export function isDate(val: unknown): val is Date {
  // @ts-ignore
  return typeof val.getMonth === "function";
}
