import { splitSheetAddress } from "../excel/address";

export type ReplySegment =
  | { readonly kind: "text"; readonly text: string }
  | { readonly kind: "ref"; readonly text: string; readonly address: string };

const CELLS = /^\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?$/;
const BACKTICKED = /`([^`\n]+)`/g;

const isCellAddress = (candidate: string): boolean => CELLS.test(splitSheetAddress(candidate).cells);

/**
 * Split a model reply into text and clickable cell references. Only backtick-quoted, row-numbered
 * addresses become refs ("`D2:D20`"), so words like Q3 or FY26 are never turned into links.
 */
export const linkifyReply = (text: string): readonly ReplySegment[] => {
  const segments: ReplySegment[] = [];
  let last = 0;
  const push = (end: number) => {
    if (end > last) segments.push({ kind: "text", text: text.slice(last, end) });
  };
  for (const match of text.matchAll(BACKTICKED)) {
    const inner = match[1] ?? "";
    if (!isCellAddress(inner)) continue;
    const start = match.index ?? 0;
    push(start);
    segments.push({ kind: "ref", text: inner, address: inner });
    last = start + match[0].length;
  }
  push(text.length);
  return segments;
};
