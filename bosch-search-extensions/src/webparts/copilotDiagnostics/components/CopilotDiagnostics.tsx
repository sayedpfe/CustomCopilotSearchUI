import * as React from 'react';
import { useState, useRef } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useGraphClient } from '../../../hooks/useGraphClient';
import { CopilotChatService, IStreamCallbacks } from '../../../services/CopilotChatService';
import styles from './CopilotDiagnostics.module.scss';

type EndpointMode = 'chat' | 'chatOverStream';
type GroundingMode = 'work' | 'web' | 'both';

interface IChunkEvent {
  offsetMs: number;   // ms since test start
  textLength: number;
  delta: number;      // chars added since last chunk
}

interface ITestRun {
  id: number;
  mode: EndpointMode;
  question: string;
  startedAt: Date;
  // Timing (all from test start, so directly comparable)
  createConversationMs: number | null;
  firstTokenMs: number | null;  // chatOverStream: first chunk with text
  completeMs: number | null;    // chat: response received; chatOverStream: onDone fired
  totalMs: number | null;
  // Content
  responseText: string;
  wordCount: number;
  chunkEvents: IChunkEvent[];
  error: string | null;
}

export interface ICopilotDiagnosticsProps {
  context: WebPartContext;
  endpointMode: EndpointMode;
  groundingMode: GroundingMode;
  defaultQuestion: string;
}

export const CopilotDiagnostics: React.FC<ICopilotDiagnosticsProps> = ({
  context,
  endpointMode: defaultMode,
  groundingMode: defaultGrounding,
  defaultQuestion,
}) => {
  const { graphClient } = useGraphClient(context);
  const [question, setQuestion] = useState(defaultQuestion || 'What are the key company priorities for this year?');
  const [mode, setMode] = useState<EndpointMode>(defaultMode || 'chatOverStream');
  const [grounding, setGrounding] = useState<GroundingMode>(defaultGrounding || 'work');
  const [isRunning, setIsRunning] = useState(false);
  const [runs, setRuns] = useState<ITestRun[]>([]);
  const [expandedRun, setExpandedRun] = useState<number | null>(null);
  const runIdRef = useRef(0);

  const formatMs = (ms: number | null): string => {
    if (ms === null) return '—';
    if (ms < 1000) return `${ms}ms`;
    return `${(ms / 1000).toFixed(2)}s`;
  };

  const runTest = async (): Promise<void> => {
    if (!graphClient || !question.trim() || isRunning) return;

    const runId = ++runIdRef.current;
    const t0 = performance.now();

    const newRun: ITestRun = {
      id: runId,
      mode,
      question: question.trim(),
      startedAt: new Date(),
      createConversationMs: null,
      firstTokenMs: null,
      completeMs: null,
      totalMs: null,
      responseText: '',
      wordCount: 0,
      chunkEvents: [],
      error: null,
    };

    setRuns((prev) => [newRun, ...prev]);
    setExpandedRun(runId);
    setIsRunning(true);

    const update = (patch: Partial<ITestRun>): void => {
      setRuns((prev) => prev.map((r) => (r.id === runId ? { ...r, ...patch } : r)));
    };

    const svc = new CopilotChatService(graphClient);

    try {
      // Step 1: Create conversation — common for both modes
      const conv = await svc.createConversation();
      const createConversationMs = Math.round(performance.now() - t0);
      update({ createConversationMs });

      if (mode === 'chat') {
        // ── Synchronous chat ──────────────────────────────────────────────────
        const result = await svc.sendMessage(conv.id, question.trim(), {
          enableWebGrounding: grounding !== 'work',
        });
        const completeMs = Math.round(performance.now() - t0);
        update({
          completeMs,
          totalMs: completeMs,
          responseText: result.responseText,
          wordCount: result.responseText.split(/\s+/).filter(Boolean).length,
        });

      } else {
        // ── SSE streaming ─────────────────────────────────────────────────────
        const chunkEvents: IChunkEvent[] = [];
        let firstTokenMs: number | null = null;
        let lastLength = 0;

        const callbacks: IStreamCallbacks = {
          onChunk: (textSoFar) => {
            const now = Math.round(performance.now() - t0);
            if (firstTokenMs === null && textSoFar.length > 0) {
              firstTokenMs = now;
            }
            const delta = textSoFar.length - lastLength;
            lastLength = textSoFar.length;
            chunkEvents.push({ offsetMs: now, textLength: textSoFar.length, delta });
            update({
              firstTokenMs,
              chunkEvents: [...chunkEvents],
              responseText: textSoFar,
            });
          },
          onDone: (fullText) => {
            const completeMs = Math.round(performance.now() - t0);
            update({
              completeMs,
              totalMs: completeMs,
              responseText: fullText,
              wordCount: fullText.split(/\s+/).filter(Boolean).length,
              chunkEvents: [...chunkEvents],
            });
          },
          onError: (err) => {
            update({ error: err.message });
          },
        };

        await svc.sendMessageStream(conv.id, question.trim(), callbacks, {
          enableWebGrounding: grounding !== 'work',
        });
      }
    } catch (err) {
      update({ error: (err as Error).message || String(err) });
    } finally {
      setIsRunning(false);
    }
  };

  // Diagnosis: are SSE chunks truly incremental or arriving in one burst?
  const getChunkDiagnosis = (events: IChunkEvent[]): { isBursty: boolean; spreadMs: number } => {
    if (events.length < 2) return { isBursty: false, spreadMs: 0 };
    const spreadMs = events[events.length - 1].offsetMs - events[0].offsetMs;
    // "Bursty" = all chunks arrived within 800ms despite having many events
    const isBursty = spreadMs < 800 && events.length >= 3;
    return { isBursty, spreadMs };
  };

  return (
    <div className={styles.container}>
      {/* Header */}
      <div className={styles.header}>
        <Icon iconName="SpeedHigh" className={styles.headerIcon} />
        <div>
          <h2 className={styles.title}>Copilot Chat Diagnostics</h2>
          <p className={styles.subtitle}>
            Compare <code>/chat</code> (sync) vs <code>/chatOverStream</code> (SSE) — measure TTFT, detect proxy buffering
          </p>
        </div>
      </div>

      {/* Config row */}
      <div className={styles.configRow}>
        <div className={styles.configField}>
          <label className={styles.label}>Endpoint mode</label>
          <div className={styles.modeToggle}>
            <button
              className={`${styles.modeBtn} ${mode === 'chat' ? styles.modeBtnActive : ''}`}
              onClick={() => setMode('chat')}
            >
              /chat  (sync)
            </button>
            <button
              className={`${styles.modeBtn} ${mode === 'chatOverStream' ? styles.modeBtnActive : ''}`}
              onClick={() => setMode('chatOverStream')}
            >
              /chatOverStream  (SSE)
            </button>
          </div>
        </div>
        <div className={styles.configField}>
          <label className={styles.label}>Grounding</label>
          <select
            className={styles.select}
            value={grounding}
            onChange={(e) => setGrounding(e.target.value as GroundingMode)}
          >
            <option value="work">Work only</option>
            <option value="web">Web only</option>
            <option value="both">Work + Web</option>
          </select>
        </div>
      </div>

      {/* Question + Run */}
      <div className={styles.inputRow}>
        <input
          className={styles.questionInput}
          type="text"
          value={question}
          onChange={(e) => setQuestion(e.target.value)}
          placeholder="Enter test question…"
          onKeyDown={(e) => { if (e.key === 'Enter') runTest(); }}
          disabled={isRunning}
        />
        <button
          className={styles.runBtn}
          onClick={runTest}
          disabled={isRunning || !graphClient || !question.trim()}
        >
          {isRunning
            ? <><Spinner className={styles.btnSpinner} /> Running…</>
            : <><Icon iconName="Play" /> Run Test</>
          }
        </button>
      </div>

      {/* Performance tips */}
      <div className={styles.tipsBox}>
        <Icon iconName="Info" className={styles.tipsIcon} />
        <div>
          <strong>How to interpret results: </strong>
          If <code>/chatOverStream</code> shows all SSE chunks arriving within ≤800ms of each other, the
          SharePoint HTTP proxy is <strong>buffering the entire response</strong> before forwarding it —
          meaning streaming gives no UX advantage and <code>/chat</code> is simpler.
          If chunks have a spread &gt;1s, streaming is genuinely incremental.
          <br />
          <strong>TTFT</strong> (Time To First Token) is the most user-visible metric — it determines
          when text starts appearing. For SSE: TTFT = first chunk offset. For sync: TTFT = completeMs.
        </div>
      </div>

      {/* Results */}
      {runs.length > 0 && (
        <div className={styles.results}>
          <h3 className={styles.resultsTitle}>Test Runs ({runs.length})</h3>
          {runs.map((run) => {
            const { isBursty, spreadMs } = getChunkDiagnosis(run.chunkEvents);
            const isExpanded = expandedRun === run.id;
            return (
              <div key={run.id} className={`${styles.runCard} ${run.error ? styles.runCardError : ''}`}>
                {/* Run header (click to expand) */}
                <div
                  className={styles.runHeader}
                  onClick={() => setExpandedRun(isExpanded ? null : run.id)}
                >
                  <span className={`${styles.runBadge} ${run.mode === 'chat' ? styles.runBadgeSync : styles.runBadgeStream}`}>
                    {run.mode === 'chat' ? '/chat' : '/chatOverStream'}
                  </span>
                  <span className={styles.runQuestion}>{run.question}</span>
                  <span className={styles.runTime}>{run.startedAt.toLocaleTimeString()}</span>
                  {run.totalMs === null && !run.error
                    ? <Spinner className={styles.runSpinner} />
                    : (
                      <span className={`${styles.runTotal} ${run.totalMs !== null && run.totalMs < 8000 ? styles.runTotalFast : ''}`}>
                        {run.error ? '❌ Error' : `⏱ ${formatMs(run.totalMs)}`}
                      </span>
                    )
                  }
                  <Icon iconName={isExpanded ? 'ChevronUp' : 'ChevronDown'} className={styles.chevron} />
                </div>

                {/* Detail panel */}
                {isExpanded && (
                  <div className={styles.runDetail}>

                    {run.error && (
                      <div className={styles.errorBox}>
                        <Icon iconName="ErrorBadge" style={{ marginRight: 8 }} />
                        {run.error}
                      </div>
                    )}

                    {/* Timing table */}
                    <table className={styles.timingTable}>
                      <thead>
                        <tr>
                          <th>Step</th>
                          <th>Time from start</th>
                          <th>Endpoint called</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td>Create conversation</td>
                          <td className={styles.timingValue}>{formatMs(run.createConversationMs)}</td>
                          <td className={styles.timingNote}>POST /beta/copilot/conversations</td>
                        </tr>
                        {run.mode === 'chatOverStream' && (
                          <tr>
                            <td>
                              <strong>First token (TTFT)</strong>
                              <span className={styles.timingKey}> ← most important</span>
                            </td>
                            <td className={`${styles.timingValue} ${run.firstTokenMs !== null ? styles.timingValueHighlight : ''}`}>
                              {formatMs(run.firstTokenMs)}
                            </td>
                            <td className={styles.timingNote}>First SSE chunk with text received</td>
                          </tr>
                        )}
                        <tr>
                          <td>Response complete</td>
                          <td className={styles.timingValue}>{formatMs(run.completeMs)}</td>
                          <td className={styles.timingNote}>
                            {run.mode === 'chat' ? 'POST /chat resolved' : 'onDone callback fired'}
                          </td>
                        </tr>
                        <tr className={styles.timingRowTotal}>
                          <td><strong>Total</strong></td>
                          <td className={styles.timingValue}><strong>{formatMs(run.totalMs)}</strong></td>
                          <td className={styles.timingNote}>
                            {run.wordCount > 0 ? `${run.wordCount} words in response` : ''}
                          </td>
                        </tr>
                      </tbody>
                    </table>

                    {/* SSE chunk analysis — streaming only */}
                    {run.mode === 'chatOverStream' && (
                      <div className={styles.chunkSection}>
                        <div className={styles.chunkHeader}>
                          SSE Chunk Timeline
                          {run.chunkEvents.length > 0 && (
                            <span className={styles.chunkCount}> ({run.chunkEvents.length} events)</span>
                          )}
                        </div>

                        {run.chunkEvents.length === 0 && run.totalMs === null && (
                          <div className={styles.chunkWaiting}>Waiting for stream…</div>
                        )}

                        {run.chunkEvents.length >= 2 && (
                          <div className={`${styles.diagnosis} ${isBursty ? styles.diagnosisBad : styles.diagnosisGood}`}>
                            {isBursty
                              ? `⚠ All ${run.chunkEvents.length} chunks arrived within ${spreadMs}ms — response appears BUFFERED. Streaming provides no UX benefit here; consider switching to /chat.`
                              : `✓ Chunks spread over ${spreadMs}ms — stream is genuinely incremental. TTFT = ${formatMs(run.firstTokenMs)}.`
                            }
                          </div>
                        )}

                        {run.chunkEvents.length > 0 && (
                          <div className={styles.chunkTimeline}>
                            {run.chunkEvents.slice(0, 25).map((ev, i) => (
                              <span
                                key={i}
                                className={styles.chunkPill}
                                title={`Chunk ${i + 1}: +${ev.offsetMs}ms | total ${ev.textLength} chars | Δ+${ev.delta} chars`}
                              >
                                +{ev.offsetMs}ms
                              </span>
                            ))}
                            {run.chunkEvents.length > 25 && (
                              <span className={styles.chunkPillMore}>+{run.chunkEvents.length - 25} more</span>
                            )}
                          </div>
                        )}
                      </div>
                    )}

                    {/* Response preview */}
                    {run.responseText && (
                      <div className={styles.responseSection}>
                        <div className={styles.responseSectionTitle}>
                          Response preview ({run.wordCount} words)
                        </div>
                        <div className={styles.responseText}>{run.responseText}</div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};
