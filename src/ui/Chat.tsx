import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, Spinner, Text, Textarea } from "@fluentui/react-components";
import { createGeminiClient } from "../agent/geminiClient";
import { trimHistory } from "../agent/history";
import { runAgent, type AgentEvent, type ChatMessage } from "../agent/loop";
import { SYSTEM_PROMPT } from "../agent/prompt";
import { createExcelExecutor } from "../excel/executors";
import { buildLiveContext } from "../excel/liveContext";
import { browserChatStore, type Line } from "../lib/chatStore";
import { copyText } from "../lib/clipboard";
import { friendlyError } from "../lib/friendlyError";
import { summarizeRange } from "../excel/context";
import { isExcelHost, readSelection } from "../excel/readSelection";
import { DEFAULT_SUGGESTIONS, suggestPrompts } from "../excel/suggestions";
import { createUndoLog } from "../excel/undo";

type NewLine = Line extends infer L ? (L extends Line ? Omit<L, "id"> : never) : never;

interface Props {
  readonly apiKey: string;
  readonly modelId: string;
}

const KEEP_USER_TURNS = 4;

const colorFor = (line: Line): string => {
  switch (line.kind) {
    case "user": return "#e8f1fb";
    case "error": return "#fde7e9";
    case "action": return line.ok ? "#eef7ee" : "#fff4ce";
    case "note": return "#f3f2f1";
    default: return "#ffffff";
  }
};

const workbookKeyOf = (): string => {
  try {
    return Office.context.document.url || "default";
  } catch {
    return "default";
  }
};

export const Chat = ({ apiKey, modelId }: Props) => {
  const chatStore = useMemo(() => browserChatStore(), []);
  const workbookKey = useMemo(workbookKeyOf, []);
  const [lines, setLines] = useState<readonly Line[]>(() => chatStore.load(workbookKey));
  const [copiedId, setCopiedId] = useState<number | null>(null);
  const lastCommandRef = useRef([...lines].reverse().find((l) => l.kind === "user")?.text ?? "");
  const [input, setInput] = useState("");
  const [running, setRunning] = useState(false);
  const [canUndo, setCanUndo] = useState(false);
  const [suggestions, setSuggestions] = useState<readonly string[]>(DEFAULT_SUGGESTIONS);
  const historyRef = useRef<readonly ChatMessage[]>([]);
  const abortRef = useRef<AbortController | null>(null);
  const idRef = useRef(Math.max(0, ...chatStore.load(workbookKey).map((l) => l.id)));
  const bottomRef = useRef<HTMLDivElement>(null);

  const undo = useMemo(() => createUndoLog(), []);
  const execute = useMemo(() => createExcelExecutor(undo), [undo]);
  const model = useMemo(
    () => createGeminiClient({ apiKey, model: modelId, systemInstruction: SYSTEM_PROMPT }),
    [apiKey, modelId],
  );

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ block: "end" });
  }, [lines]);

  useEffect(() => {
    chatStore.save(workbookKey, lines);
  }, [chatStore, workbookKey, lines]);

  // Refresh selection-aware chips whenever the chat is empty (mount, after Clear chat).
  const isEmpty = lines.length === 0;
  useEffect(() => {
    if (!isEmpty || !isExcelHost()) return;
    let cancelled = false;
    readSelection()
      .then((snap) => !cancelled && setSuggestions(suggestPrompts(summarizeRange(snap))))
      .catch(() => !cancelled && setSuggestions(DEFAULT_SUGGESTIONS));
    return () => {
      cancelled = true;
    };
  }, [isEmpty]);

  const push = useCallback((line: NewLine) => {
    idRef.current += 1;
    const next = { ...line, id: idRef.current } as Line;
    setLines((prev) => [...prev, next]);
  }, []);

  const onEvent = useCallback(
    (e: AgentEvent) => {
      if (e.type === "text") push({ kind: "assistant", text: e.text });
      else push({ kind: "action", ok: e.ok, text: `${e.ok ? "✓" : "⚠"} ${e.summary}` });
    },
    [push],
  );

  const send = useCallback(
    async (raw: string) => {
      const text = raw.trim();
      if (!text || running) return;
      if (!isExcelHost()) {
        push({ kind: "error", text: "Not running inside Excel. Open this pane from the Excel ribbon." });
        return;
      }
      setInput("");
      lastCommandRef.current = text;
      push({ kind: "user", text });
      setRunning(true);
      undo.beginGroup(text.slice(0, 60));
      const controller = new AbortController();
      abortRef.current = controller;
      try {
        const context = await buildLiveContext();
        const result = await runAgent({
          model,
          execute,
          userMessage: text,
          context,
          history: historyRef.current,
          signal: controller.signal,
          onEvent,
        });
        historyRef.current = trimHistory(result.history, KEEP_USER_TURNS);
        if (result.aborted) push({ kind: "note", text: "Stopped." });
        else if (result.hitStepLimit) push({ kind: "note", text: "Stopped after the step limit. Ask me to continue if needed." });
      } catch (err) {
        if (controller.signal.aborted) push({ kind: "note", text: "Stopped." });
        else push({ kind: "error", text: friendlyError(err) });
      } finally {
        abortRef.current = null;
        setRunning(false);
        setCanUndo(undo.canUndo());
      }
    },
    [running, push, undo, model, execute, onEvent],
  );

  const doUndo = useCallback(async () => {
    const outcome = await undo.undoLast();
    setCanUndo(undo.canUndo());
    if (!outcome) return;
    const failed = outcome.failed.length > 0 ? ` (${outcome.failed.length} could not be restored)` : "";
    push({ kind: "note", text: `Undid "${outcome.label}": ${outcome.restored} step(s)${failed}.` });
  }, [undo, push]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8, minHeight: 0, flex: 1 }}>
      <div style={{ flex: 1, overflowY: "auto", display: "grid", gap: 6, alignContent: "start", minHeight: 120 }}>
        {lines.length === 0 && (
          <div style={{ display: "grid", gap: 6 }}>
            <Text size={200}>Tell me what to do with your sheet. Try:</Text>
            {suggestions.map((ex) => (
              <Button key={ex} size="small" appearance="outline" style={{ justifyContent: "flex-start", textAlign: "left", height: "auto", padding: "4px 8px" }} onClick={() => setInput(ex)}>
                {ex}
              </Button>
            ))}
          </div>
        )}
        {lines.map((l) => (
          <div key={l.id} style={{ background: colorFor(l), borderRadius: 6, padding: "6px 8px", fontSize: 13, whiteSpace: "pre-wrap", overflowWrap: "anywhere" }}>
            {l.text}
            {(l.kind === "assistant" || l.kind === "error") && (
              <div style={{ textAlign: "right" }}>
                <Button
                  size="small"
                  appearance="transparent"
                  onClick={() => void copyText(l.text).then((done) => done && setCopiedId(l.id))}
                >
                  {copiedId === l.id ? "Copied" : "Copy"}
                </Button>
              </div>
            )}
          </div>
        ))}
        {running && <Spinner size="tiny" label="Working..." labelPosition="after" />}
        <div ref={bottomRef} />
      </div>

      <div style={{ display: "flex", gap: 6 }}>
        <Button size="small" onClick={() => void doUndo()} disabled={!canUndo || running}>Undo last</Button>
        <Button size="small" onClick={() => void send(lastCommandRef.current)} disabled={running || !lastCommandRef.current}>
          Retry last
        </Button>
        <Button size="small" appearance="subtle" onClick={() => { setLines([]); historyRef.current = []; }} disabled={running}>
          Clear chat
        </Button>
      </div>

      <Textarea
        value={input}
        onChange={(_, d) => setInput(d.value)}
        placeholder="e.g. Add a column that totals each row"
        resize="vertical"
        rows={3}
        disabled={running}
        onKeyDown={(e) => {
          if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            void send(input);
          }
        }}
      />
      {running ? (
        <Button onClick={() => abortRef.current?.abort()}>Stop</Button>
      ) : (
        <Button appearance="primary" disabled={!input.trim()} onClick={() => void send(input)}>Send</Button>
      )}
    </div>
  );
};
