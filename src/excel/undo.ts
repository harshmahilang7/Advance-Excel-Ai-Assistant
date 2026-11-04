export type Restore = () => Promise<void>;

export interface UndoOutcome {
  readonly label: string;
  readonly restored: number;
  readonly failed: readonly string[];
}

export interface UndoLog {
  /** Start recording a new command. Empty groups are dropped. */
  beginGroup(label: string): void;
  record(description: string, restore: Restore): void;
  canUndo(): boolean;
  undoLast(): Promise<UndoOutcome | null>;
}

interface Entry {
  readonly description: string;
  readonly restore: Restore;
}
interface Group {
  readonly label: string;
  readonly entries: readonly Entry[];
}

const DEFAULT_MAX_GROUPS = 20;

const messageOf = (err: unknown): string => (err instanceof Error ? err.message : String(err));

/** Per-command undo history. Groups are replaced, never mutated in place. */
export const createUndoLog = (maxGroups: number = DEFAULT_MAX_GROUPS): UndoLog => {
  let groups: readonly Group[] = [];
  let open: Group | null = null;

  const closeOpen = () => {
    if (open && open.entries.length > 0) groups = [...groups, open].slice(-maxGroups);
    open = null;
  };

  return {
    beginGroup(label) {
      closeOpen();
      open = { label, entries: [] };
    },
    record(description, restore) {
      if (!open) open = { label: "command", entries: [] };
      open = { ...open, entries: [...open.entries, { description, restore }] };
    },
    canUndo() {
      return groups.length > 0 || (open?.entries.length ?? 0) > 0;
    },
    async undoLast() {
      closeOpen();
      const group = groups.at(-1);
      if (!group) return null;
      groups = groups.slice(0, -1);
      const failed: string[] = [];
      let restored = 0;
      for (const entry of [...group.entries].reverse()) {
        try {
          await entry.restore();
          restored += 1;
        } catch (err) {
          failed.push(`${entry.description}: ${messageOf(err)}`);
        }
      }
      return { label: group.label, restored, failed };
    },
  };
};
