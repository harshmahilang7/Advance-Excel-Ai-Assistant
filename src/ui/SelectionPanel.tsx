import { useCallback, useState } from "react";
import { Button, MessageBar, MessageBarBody, Text } from "@fluentui/react-components";
import { summarizeRange, type SheetSummary } from "../excel/context";
import { isExcelHost, readSelection } from "../excel/readSelection";

export const SelectionPanel = () => {
  const [summary, setSummary] = useState<SheetSummary | null>(null);
  const [error, setError] = useState<string | null>(null);

  const refresh = useCallback(async () => {
    setError(null);
    try {
      if (!isExcelHost()) throw new Error("Not running inside Excel. Sideload the add-in to test.");
      setSummary(summarizeRange(await readSelection()));
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not read the selection.");
    }
  }, []);

  return (
    <div style={{ display: "grid", gap: 8 }}>
      <Button onClick={() => void refresh()}>Read my selection</Button>
      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
      {summary && (
        <pre style={{ fontSize: 11, whiteSpace: "pre-wrap", margin: 0 }}>
          {JSON.stringify(summary, null, 2)}
        </pre>
      )}
      {!summary && !error && <Text size={200}>Select cells in Excel, then click above.</Text>}
    </div>
  );
};
