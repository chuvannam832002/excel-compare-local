import { useState, type KeyboardEvent } from 'react';
import { Link } from 'wouter';
import {
  ArrowLeft, Check, CircleAlert, Clock3, Code2, Eraser, LoaderCircle,
  LockKeyhole, Play, RotateCcw, ShieldCheck, Terminal,
} from 'lucide-react';
import { runPython, type PythonRunResult, type PythonRunStatus } from '@/lib/python-runtime';

const SAMPLE_CODE = `# A small, self-contained Python example
prices = [18.50, 24.00, 16.75, 31.20]
average = sum(prices) / len(prices)

print(f"Items: {len(prices)}")
print(f"Average price: \${average:.2f}")
print(f"Highest price: \${max(prices):.2f}")

average`;

type ViewStatus = 'idle' | PythonRunStatus | 'complete' | 'error' | 'timeout';

function PythonRunner() {
  const [code, setCode] = useState(SAMPLE_CODE);
  const [status, setStatus] = useState<ViewStatus>('idle');
  const [runResult, setRunResult] = useState<PythonRunResult | null>(null);
  const [errorMessage, setErrorMessage] = useState('');

  const clearOutput = () => {
    setRunResult(null);
    setErrorMessage('');
    setStatus('idle');
  };

  const run = async () => {
    if (status === 'loading' || status === 'running') return;
    setRunResult(null);
    setErrorMessage('');
    setStatus('loading');
    try {
      const result = await runPython(code, (nextStatus: PythonRunStatus) => setStatus(nextStatus));
      setRunResult(result);
      if (result.error) {
        const timedOut = /timed?\s*out|timeout|time limit/i.test(result.error);
        setStatus(timedOut ? 'timeout' : 'error');
        setErrorMessage(result.error);
      } else {
        setStatus('complete');
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Python could not complete this run.';
      const timedOut = /timed?\s*out|timeout|time limit/i.test(message);
      setErrorMessage(message);
      setStatus(timedOut ? 'timeout' : 'error');
    }
  };

  const onEditorKeyDown = (event: KeyboardEvent<HTMLTextAreaElement>) => {
    if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') {
      event.preventDefault();
      void run();
    }
  };

  const isBusy = status === 'loading' || status === 'running';
  const statusLabel = status === 'loading'
    ? 'Starting Python…'
    : status === 'running'
      ? 'Running in your browser…'
      : status === 'complete'
        ? 'Run complete'
        : status === 'timeout'
          ? 'Time limit reached'
          : status === 'error'
            ? 'Run encountered an error'
            : 'Ready when you are';

  return (
    <div className="python-page">
      <header className="python-topbar">
        <Link href="/" className="python-brand" data-testid="link-back-compare" aria-label="Back to Sheetwise comparison">
          <span className="python-brand-mark"><Code2 size={17} strokeWidth={2.2} /></span>
          <span className="python-brand-name">sheetwise<span> / python</span></span>
        </Link>
        <div className="python-topbar-right">
          <span className="python-local-badge"><ShieldCheck size={14} /> Runs on this device</span>
          <Link href="/" className="python-back-link" data-testid="link-back-workspace">
            <ArrowLeft size={15} /> Spreadsheet compare
          </Link>
        </div>
      </header>

      <main className="python-workspace">
        <section className="python-intro">
          <div className="python-intro-copy">
            <div className="python-eyebrow"><span /> BROWSER PYTHON WORKSPACE</div>
            <h1>A little room to <em>think in code.</em></h1>
            <p>Try a short Python script and inspect what it prints. Your code stays in this browser, beside your spreadsheet work.</p>
          </div>
          <aside className="python-privacy-card" aria-label="Privacy information">
            <span className="python-privacy-icon"><LockKeyhole size={17} /></span>
            <div><strong>Private by design</strong><span>The Python runtime downloads from the internet on first use. Your script runs locally and is not sent to a server.</span></div>
          </aside>
        </section>

        <section className="python-runner" aria-label="Python code runner">
          <div className="python-editor-column">
            <div className="python-panel-heading">
              <div className="python-heading-title"><span className="python-section-mark"><Code2 size={15} /></span><div><strong>Script</strong><span>Write or paste a short Python snippet</span></div></div>
              <span className="python-language-tag">PYTHON 3</span>
            </div>
            <div className="python-editor-shell">
              <div className="python-editor-gutter" aria-hidden="true">
                {code.split('\n').map((_, index) => <span key={index}>{index + 1}</span>)}
              </div>
              <textarea
                className="python-code-editor"
                value={code}
                onChange={(event) => setCode(event.target.value)}
                onKeyDown={onEditorKeyDown}
                spellCheck={false}
                aria-label="Python script editor"
                data-testid="input-python-code"
                autoCapitalize="off"
                autoComplete="off"
                autoCorrect="off"
              />
            </div>
            <div className="python-editor-footer">
              <span className="python-shortcut"><kbd>Ctrl</kbd><span>+</span><kbd>Enter</kbd> to run</span>
              <span>{code.length.toLocaleString()} characters</span>
            </div>
            <div className="python-action-bar">
              <div className="python-action-secondary">
                <button type="button" className="python-tool-button" onClick={() => { setCode(''); clearOutput(); }} disabled={isBusy} data-testid="button-clear-code" aria-label="Clear script">
                  <Eraser size={15} /> Clear
                </button>
                <button type="button" className="python-tool-button" onClick={() => { setCode(SAMPLE_CODE); clearOutput(); }} disabled={isBusy} data-testid="button-reset-sample">
                  <RotateCcw size={14} /> Reset sample
                </button>
              </div>
              <button type="button" className="python-run-button" onClick={() => void run()} disabled={isBusy} data-testid="button-run-python">
                {isBusy ? <LoaderCircle size={16} className="python-spinning" /> : <Play size={15} fill="currentColor" />}
                {status === 'loading' ? 'Starting…' : status === 'running' ? 'Running…' : 'Run script'}
              </button>
            </div>
          </div>

          <section className={`python-output-column ${status === 'error' || status === 'timeout' ? 'has-error' : ''}`} aria-label="Python output" aria-live="polite">
            <div className="python-output-heading">
              <div className="python-heading-title"><span className="python-section-mark output-mark"><Terminal size={15} /></span><div><strong>Output</strong><span>Printed text and returned value</span></div></div>
              <button type="button" className="python-clear-output" onClick={clearOutput} data-testid="button-clear-output" disabled={isBusy || (!runResult && !errorMessage && status === 'idle')} aria-label="Clear output">
                <Eraser size={14} /> Clear
              </button>
            </div>

            <div className={`python-run-status status-${status}`} data-testid="status-python-run">
              <span className="python-status-indicator">{isBusy ? <LoaderCircle size={13} className="python-spinning" /> : status === 'complete' ? <Check size={13} /> : status === 'error' || status === 'timeout' ? <CircleAlert size={13} /> : <span />}</span>
              <span>{statusLabel}</span>
              {runResult && <span className="python-duration"><Clock3 size={12} /> {runResult.durationMs} ms</span>}
            </div>

            {errorMessage && (
              <div className={`python-error-callout ${status === 'timeout' ? 'timeout' : ''}`} role="alert" data-testid="error-python-run">
                <strong>{status === 'timeout' ? 'This run reached its time limit.' : 'Python reported an error.'}</strong>
                <pre>{errorMessage}</pre>
              </div>
            )}

            <div className="python-output-block" data-testid="output-stdout">
              <div className="python-output-label"><span className="python-output-dot stdout-dot" /> STDOUT <span>printed text</span></div>
              <pre>{runResult?.stdout || <span className="python-output-placeholder">{status === 'idle' ? 'Printed lines will appear here.' : 'No printed output.'}</span>}</pre>
            </div>
            {runResult?.stderr && (
              <div className="python-output-block stderr-block" data-testid="output-stderr">
                <div className="python-output-label"><span className="python-output-dot stderr-dot" /> STDERR <span>diagnostics</span></div>
                <pre>{runResult.stderr}</pre>
              </div>
            )}
            <div className="python-output-block result-block" data-testid="output-result">
              <div className="python-output-label"><span className="python-output-dot result-dot" /> RESULT <span>last expression</span></div>
              <pre>{runResult?.result || <span className="python-output-placeholder">{status === 'idle' ? 'A returned value will appear here.' : 'No returned value.'}</span>}</pre>
            </div>
            {!runResult && !errorMessage && status === 'idle' && (
              <div className="python-output-empty" data-testid="empty-python-output">
                <span><Terminal size={17} /></span>
                <p>Run your script to see what it does.</p>
              </div>
            )}
          </section>
        </section>

        <footer className="python-footer">
          <span><LockKeyhole size={12} /> Your code is temporary and stays in this browser.</span>
          <span className="python-footer-divider" />
          <span>For quick calculations and small experiments.</span>
        </footer>
      </main>
    </div>
  );
}

export default PythonRunner;
