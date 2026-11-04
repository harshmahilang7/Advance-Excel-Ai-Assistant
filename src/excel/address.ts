export interface SplitAddress {
  /** Unquoted sheet name, or null when the address is not sheet-qualified. */
  readonly sheet: string | null;
  readonly cells: string;
}

const QUALIFIED = /^(?:'((?:[^']|'')+)'|([^!']+))!(.+)$/;

/** "'Bob''s Data'!B2" -> { sheet: "Bob's Data", cells: "B2" } */
export const splitSheetAddress = (address: string): SplitAddress => {
  const match = QUALIFIED.exec(address);
  if (!match) return { sheet: null, cells: address };
  const quoted = match[1];
  const plain = match[2];
  const cells = match[3] ?? address;
  return { sheet: quoted !== undefined ? quoted.replace(/''/g, "'") : (plain ?? null), cells };
};
