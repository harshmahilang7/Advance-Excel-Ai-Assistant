import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, Spinner, Switch, Text, Textarea } from "@fluentui/react-components";
import { addEntry, dropLastMatching, formatChangeLog, type ChangeEntry } from "../agent/changeLog";
import { applyPending, createPreviewExecutor, isReadOnlyTool, type PendingCall } from "../agent/preview";
import { isPlanTool } from "../agent/plan";
import { ReplyText } from "./ReplyText";
import { importTable } from "../excel/importTable";
import { parseDelimited } from "../lib/csv";
import { addTemplate, browserTemplateStore, removeTemplate, type Template } from "../lib/templateStore";
import { formatRunSummary } from "../lib/runSummary";
import { addRead, describeRead, formatReadLog } from "../agent/readLog";
import { browserChangeLogStore } from "../lib/changeLogStore";
import { createGeminiClient } from "../agent/geminiClient";
import { trimHistory } from "../agent/history";
import { runAgent, type AgentEvent, type ChatMessage, type StreamEvent } from "../agent/loop";
import { CONTINUE_PROMPT, SYSTEM_PROMPT } from "../agent/prompt";
import { estimateHistoryTokens, estimateTokens, formatTokens, LARGE_CONTEXT_TOKENS } from "../agent/tokens";
import { toolDeclarations } from "../agent/tools";
import { createExcelExecutor } from "../excel/executors";
import { buildLiveContext } from "../excel/liveContext";
import { browserChatStore, type Line } from "../lib/chatStore";
import { copyText } from "../lib/clipboard";
import { friendlyError } from "../lib/friendlyError";
import { summarizeRange } from "../excel/context";
import { isExcelHost, readSelection } from "../excel/readSelection";
import { DEFAULT_SUGGESTIONS, REVIEW_SUGGESTIONS, suggestPrompts } from "../excel/suggestions";
import { createUndoLog } from "../excel/undo";

type NewLine = Line extends infer L ? (L extends Line ? Omit<L, "id"> : never) : never;

interface Props {
  readonly apiKeys: readonly string[];
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

const WAITING = "Waiting for the AI...";
const MAX_IMPORT_BYTES = 5_000_000;
/** Tokens every request carries before any chat: system prompt plus tool declarations. */
const FIXED_TOKENS = estimateTokens(SYSTEM_PROMPT) + estimateTokens(JSON.stringify(toolDeclarations));

const workbookKeyOf =(): string => {
  try {
    return Office.context.document.url || "default";
  } catch {
    return "default";
  }
};

export const Chat = ({ apiKeys, modelId }: Props) => {
  const chatStore = useMemo(() => browserChatStore(), []);
  const workbookKey = useMemo(workbookKeyOf, []);
  const [lines, setLines] = useState<readonly Line[]>(() => chatStore.load(workbookKey));
  const changeStore = useMemo(() => browserChangeLogStore(), []);
  const changeLogRef = useRef<readonly ChangeEntry[]>([]);
  const [copiedId, setCopiedId] = useState<number | null>(null);
  const lastCommandRef = useRef([...lines].reverse().find((l) => l.kind === "user")?.text ?? "");
  const [input, setInput] = useState("");
  const [running, setRunning] = useState(false);
  const [canUndo, setCanUndo] = useState(false);
  const [canContinue, setCanContinue] = useState(false);
  const [status, setStatus] = useState(WAITING);
  const [streamText, setStreamText] = useState("");
  const [contextTokens, setContextTokens] = useState(FIXED_TOKENS);
  const [previewMode, setPreviewMode] = useState(false);
  const [pending, setPending] = useState<readonly PendingCall[]>([]);
  const pendingLabelRef = useRef("");
  const discardNoteRef = useRef("");
  const [suggestions, setSuggestions] = useState<readonly string[]>(DEFAULT_SUGGESTIONS);
  const templateStore = useMemo(() => browserTemplateStore(), []);
  const [templates, setTemplates] = useState<readonly Template[]>(() => templateStore.load());
  const importRef = useRef<HTMLInputElement>(null);
  const historyRef = useRef<readonly ChatMessage[]>([]);
  const abortRef = useRef<AbortController | null>(null);
  const idRef = useRef(Math.max(0, ...chatStore.load(workbookKey).map((l) => l.id)));
  const bottomRef = useRef<HTMLDivElement>(null);

  const undo = useMemo(() => createUndoLog(), []);
  const execute = useMemo(() => createExcelExecutor(undo), [undo]);
  const model = useMemo(
    () => createGeminiClient({ apiKeys, model: modelId, systemInstruction: SYSTEM_PROMPT }),
    [apiKeys, modelId],
  );

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ block: "end" });
  }, [lines, streamText]);

  useEffect(() => {
    chatStore.save(workbookKey, lines);
  }, [chatStore, workbookKey, lines]);

  // Load what was changed earlier in this workbook (survives reloads and history trimming).
  useEffect(() => {
    changeLogRef.current = changeStore.load(workbookKey);
  }, [changeStore, workbookKey]);

  const setChangeLog = useCallback(
    (next: readonly ChangeEntry[]) => {
      changeLogRef.current = next;
      changeStore.save(workbookKey, next);
    },
    [changeStore, workbookKey],
  );

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
      else if (isPlanTool(e.name) && e.ok) push({ kind: "note", text: e.summary });
      else {
        push({ kind: "action", ok: e.ok, text: `${e.ok ? "✓" : "⚠"} ${e.summary}` });
        setStatus(`${WAITING} Last step: ${e.summary}`);
      }
    },
    [push],
  );

  const onStream = useCallback((e: StreamEvent) => {
    if (e.type === "delta") {
      setStreamText((prev) => prev + e.text);
      setStatus("The AI is replying...");
    } else {
      setStreamText("");
    }
  }, []);

  const send = useCallback(
    async (raw: string, keepLast = false) => {
      const text = raw.trim();
      if (!text || running) return;
      setCanContinue(false);
      if (!isExcelHost()) {
        push({ kind: "error", text: "Not running inside Excel. Open this pane from the Excel ribbon." });
        return;
      }
      setInput("");
      if (!keepLast) lastCommandRef.current = text;
      push({ kind: "user", text });
      setStatus(WAITING);
      setRunning(true);
      undo.beginGroup(text.slice(0, 60));
      const controller = new AbortController();
      abortRef.current = controller;
      let queued: readonly PendingCall[] = [];
      let steps: readonly string[] = [];
      let reads: readonly string[] = [];
      let changes = 0;
      const startedAt = Date.now();
      const collect = (e: AgentEvent) => {
        onEvent(e);
        if (e.type === "tool" && e.ok) {
          const read = describeRead(e.name, e.args);
          if (read) reads = addRead(reads, read);
          if (!isReadOnlyTool(e.name)) changes += 1;
        }
        // Only real, successful writes count. In preview mode nothing ran, so nothing is recorded here.
        if (!previewMode && e.type === "tool" && e.ok && !isReadOnlyTool(e.name)) steps = [...steps, e.summary];
      };
      const runner = previewMode
        ? createPreviewExecutor(execute, (c) => {
            queued = [...queued, c];
          })
        : execute;
      try {
        const memory = formatChangeLog(changeLogRef.current);
        const context = discardNoteRef.current + (memory ? `${memory}\n` : "") + (await buildLiveContext());
        discardNoteRef.current = "";
        const result = await runAgent({
          model,
          execute: runner,
          userMessage: text,
          context,
          history: historyRef.current,
          signal: controller.signal,
          selfCheck: !previewMode,
          onEvent: collect,
          onStream,
        });
        historyRef.current = trimHistory(result.history, KEEP_USER_TURNS);
        push({
          kind: "note",
          text: `${formatRunSummary({ ms: Date.now() - startedAt, toolCalls: result.toolCalls, changes, queued: previewMode })}
${formatReadLog(reads)}`,
        });
        if (result.aborted) push({ kind: "note", text: "Stopped." });
        else if (result.hitStepLimit) {
          push({ kind: "note", text: "Stopped after the step limit. Click Continue to keep going." });
          setCanContinue(true);
        }
      } catch (err) {
        if (controller.signal.aborted) push({ kind: "note", text: "Stopped." });
        else push({ kind: "error", text: friendlyError(err) });
      } finally {
        abortRef.current = null;
        setRunning(false);
        setCanUndo(undo.canUndo());
        setContextTokens(FIXED_TOKENS + estimateHistoryTokens(historyRef.current));
        setChangeLog(addEntry(changeLogRef.current, { command: text, steps }));
        if (queued.length > 0) {
          setPending(queued);
          pendingLabelRef.current = text.slice(0, 60);
        }
      }
    },
    [running, previewMode, push, undo, model, execute, onEvent, onStream, setChangeLog],
  );

  const saveTemplate = useCallback(() => {
    const result = addTemplate(templates, input.trim() || lastCommandRef.current);
    if (result.status === "added") {
      setTemplates(result.list);
      templateStore.save(result.list);
    }
    const message = {
      added: "Saved as a template. Find it on the start screen after Clear chat.",
      duplicate: "That command is already saved as a template.",
      empty: "Type a command, or send one first, then save it as a template.",
      full: "You have the maximum number of templates. Delete one first.",
    }[result.status];
    push({ kind: "note", text: message });
  }, [templates, templateStore, input, push]);

  const deleteTemplate = useCallback(
    (id: string) => {
      const next = removeTemplate(templates, id);
      setTemplates(next);
      templateStore.save(next);
    },
    [templates, templateStore],
  );

  const importFile = useCallback(
    async (file: File) => {
      if (running || pending.length > 0) return;
      if (!isExcelHost()) {
        push({ kind: "error", text: "Not running inside Excel. Open this pane from the Excel ribbon." });
        return;
      }
      if (file.size > MAX_IMPORT_BYTES) {
        push({ kind: "error", text: `That file is too large to import (over ${MAX_IMPORT_BYTES / 1_000_000} MB).` });
        return;
      }
      setRunning(true);
      setStatus(`Importing ${file.name}...`);
      undo.beginGroup(`import ${file.name}`.slice(0, 60));
      try {
        const table = parseDelimited(await file.text());
        if (table.rows.length === 0) throw new Error("The file has no rows to import.");
        const out = await importTable(table.rows, file.name, undo);
        const clipped = table.truncated ? " Rows or columns beyond the limits were left out." : "";
        push({ kind: "action", ok: true, text: `✓ Imported ${out.rows} rows × ${out.columns} columns into a new sheet "${out.sheet}".${clipped}` });
        setChangeLog(
          addEntry(changeLogRef.current, {
            command: `Import ${file.name}`,
            steps: [`Imported ${out.rows} rows into new sheet "${out.sheet}" (headers in row 1)`],
          }),
        );
      } catch (err) {
        push({ kind: "error", text: friendlyError(err) });
      } finally {
        setRunning(false);
        setCanUndo(undo.canUndo());
      }
    },
    [running, pending.length, push, undo, setChangeLog],
  );

  const applyChanges = useCallback(async () => {
    if (pending.length === 0 || running) return;
    setStatus(`Applying ${pending.length} change(s) to the workbook...`);
    setRunning(true);
    undo.beginGroup(pendingLabelRef.current);
    try {
      const outcome = await applyPending(execute, pending);
      for (const step of outcome.results) push({ kind: "action", ok: step.ok, text: `${step.ok ? "✓" : "⚠"} ${step.text}` });
      if (outcome.skipped > 0) push({ kind: "note", text: `Stopped after a failure; ${outcome.skipped} queued change(s) were not applied.` });
      const done = outcome.results.filter((r) => r.ok).map((r) => r.text);
      setChangeLog(addEntry(changeLogRef.current, { command: pendingLabelRef.current, steps: done }));
    } finally {
      setPending([]);
      setRunning(false);
      setCanUndo(undo.canUndo());
    }
  }, [pending, running, undo, execute, push, setChangeLog]);

  const discardChanges = useCallback(() => {
    push({ kind: "note", text: `Discarded ${pending.length} queued change(s). Nothing was changed.` });
    discardNoteRef.current = "[Note: the changes queued for the previous request were discarded by the user and NOT applied to the workbook.]\n";
    setPending([]);
  }, [pending.length, push]);

  const doUndo = useCallback(async () => {
    const outcome = await undo.undoLast();
    setCanUndo(undo.canUndo());
    if (!outcome) return;
    setChangeLog(dropLastMatching(changeLogRef.current, outcome.label));
    const failed = outcome.failed.length > 0 ? ` (${outcome.failed.length} could not be restored)` : "";
    push({ kind: "note", text: `Undid "${outcome.label}": ${outcome.restored} step(s)${failed}.` });
  }, [undo, push, setChangeLog]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 6, minHeight: 0, flex: 1, overflow: "hidden" }}>
      <div style={{ flex: "1 1 0", overflowY: "auto", display: "grid", gap: 6, alignContent: "start", minHeight: 60 }}>
        {lines.length === 0 && (
          <div style={{ display: "grid", gap: 6 }}>
            <Text size={200}>Tell me what to do with your sheet. Try:</Text>
            {suggestions.map((ex) => (
              <Button key={ex} size="small" appearance="outline" style={{ justifyContent: "flex-start", textAlign: "left", height: "auto", padding: "4px 8px" }} onClick={() => setInput(ex)}>
                {ex}
              </Button>
            ))}
            {templates.length > 0 && (
              <>
                <Text size={200} style={{ marginTop: 4 }}>Your templates:</Text>
                {templates.map((t) => (
                  <div key={t.id} style={{ display: "flex", gap: 4, alignItems: "center" }}>
                    <Button size="small" appearance="outline" title={t.prompt} style={{ flex: 1, minWidth: 0, justifyContent: "flex-start", textAlign: "left", height: "auto", padding: "4px 8px" }} onClick={() => setInput(t.prompt)}>
                      {t.name}
                    </Button>
                    <Button size="small" appearance="subtle" aria-label={`Delete template ${t.name}`} title="Delete template" onClick={() => deleteTemplate(t.id)}>
                      ✕
                    </Button>
                  </div>
                ))}
              </>
            )}
            <Text size={200} style={{ marginTop: 4 }}>Or check your sheet without changing it:</Text>
            {REVIEW_SUGGESTIONS.map((ex) => (
              <Button key={ex} size="small" appearance="subtle" style={{ justifyContent: "flex-start", textAlign: "left", height: "auto", padding: "4px 8px" }} onClick={() => setInput(ex)}>
                {ex}
              </Button>
            ))}
          </div>
        )}
        {lines.map((l) => (
          <div key={l.id} style={{ background: colorFor(l), borderRadius: 6, padding: "6px 8px", fontSize: 13, whiteSpace: "pre-wrap", overflowWrap: "anywhere" }}>
            {l.kind === "assistant" ? <ReplyText text={l.text} /> : l.text}
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
        {streamText && (
          <div style={{ background: "#ffffff", border: "1px dashed #c8c6c4", borderRadius: 6, padding: "6px 8px", fontSize: 13, whiteSpace: "pre-wrap", overflowWrap: "anywhere" }}>
            {streamText}
          </div>
        )}
        {pending.length > 0 && !running && (
          <div style={{ background: "#fff4ce", borderRadius: 6, padding: "6px 8px", fontSize: 13, display: "grid", gap: 6 }}>
            <Text weight="semibold" size={200}>Review {pending.length} queued change(s). Nothing has been applied yet.</Text>
            <ol style={{ margin: 0, paddingLeft: 18 }}>
              {pending.map((p, i) => (
                <li key={i} style={{ overflowWrap: "anywhere" }}>{p.summary}</li>
              ))}
            </ol>
            <div style={{ display: "flex", gap: 6 }}>
              <Button size="small" appearance="primary" onClick={() => void applyChanges()}>Apply</Button>
              <Button size="small" onClick={discardChanges}>Discard</Button>
            </div>
          </div>
        )}
        <div ref={bottomRef} />
      </div>

      <div style={{ display: "grid", gap: 6, flexShrink: 0 }}>
      {running && (
        <div
          role="status"
          aria-live="polite"
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            padding: "6px 8px",
            borderRadius: 6,
            background: "#e8f1fb",
            border: "1px solid #b3d1f2",
            fontSize: 13,
          }}
        >
          <Spinner size="tiny" />
          <span style={{ flex: 1, minWidth: 0, overflowWrap: "anywhere" }}>{status}</span>
        </div>
      )}

      <Text size={100} style={{ color: contextTokens >= LARGE_CONTEXT_TOKENS ? "#a4262c" : "#605e5c" }}>
        {contextTokens >= LARGE_CONTEXT_TOKENS
          ? `Context is large (~${formatTokens(contextTokens)} tokens per request). Clear chat to cut cost.`
          : `Context: ~${formatTokens(contextTokens)} tokens per request`}
      </Text>

      <Switch
        checked={previewMode}
        onChange={(_, d) => setPreviewMode(d.checked)}
        label="Preview before applying"
        disabled={running || pending.length > 0}
      />

      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", flexShrink: 0 }}>
        <Button size="small" onClick={() => void doUndo()} disabled={!canUndo || running || pending.length > 0}>Undo last</Button>
        {canContinue && (
          <Button size="small" appearance="primary" onClick={() => void send(CONTINUE_PROMPT, true)} disabled={running || pending.length > 0}>
            Continue
          </Button>
        )}
        <Button size="small" onClick={() => void send(lastCommandRef.current)} disabled={running || pending.length > 0 || !lastCommandRef.current}>
          Retry last
        </Button>
        <Button size="small" onClick={saveTemplate} disabled={running || pending.length > 0 || (!input.trim() && !lastCommandRef.current)}>
          Save template
        </Button>
        <Button size="small" onClick={() => importRef.current?.click()} disabled={running || pending.length > 0}>
          Import CSV
        </Button>
        <input
          ref={importRef}
          type="file"
          accept=".csv,.tsv,.txt,text/csv,text/plain"
          style={{ display: "none" }}
          onChange={(e) => {
            const file = e.target.files?.[0];
            e.target.value = "";
            if (file) void importFile(file);
          }}
        />
        <Button size="small" appearance="subtle" onClick={() => { setLines([]); setPending([]); setCanContinue(false); historyRef.current = []; setContextTokens(FIXED_TOKENS); changeLogRef.current = []; changeStore.clear(workbookKey); }} disabled={running}>
          Clear chat
        </Button>
      </div>

      <Textarea
        value={input}
        onChange={(_, d) => setInput(d.value)}
        placeholder="e.g. Add a column that totals each row"
        resize="none"
        rows={2}
        disabled={running || pending.length > 0}
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
        <Button appearance="primary" disabled={!input.trim() || pending.length > 0} onClick={() => void send(input)}>Send</Button>
      )}
      </div>
    </div>
  );
};
