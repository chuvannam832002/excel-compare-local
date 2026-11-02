import { loadPyodide, version } from 'pyodide';

type RunMessage = { type: 'run'; id: number; code: string };

type RunResult = {
  stdout: string;
  stderr: string;
  result: string;
  durationMs: number;
  error?: string;
};

let pyodidePromise: ReturnType<typeof loadPyodide> | null = null;

const loadRuntime = () => {
  if (!pyodidePromise) {
    pyodidePromise = loadPyodide({
      indexURL: `https://cdn.jsdelivr.net/pyodide/v${version}/full/`,
    });
  }
  return pyodidePromise;
};

self.addEventListener('message', (event: MessageEvent<RunMessage>) => {
  if (event.data.type !== 'run') return;
  void execute(event.data);
});

async function execute({ id, code }: RunMessage) {
  self.postMessage({ type: 'status', id, status: 'loading' });
  const stdout: string[] = [];
  const stderr: string[] = [];

  try {
    const pyodide = await loadRuntime();
    pyodide.setStdout({ batched: (output) => stdout.push(output) });
    pyodide.setStderr({ batched: (output) => stderr.push(output) });

    self.postMessage({ type: 'status', id, status: 'running' });
    const startedAt = performance.now();
    const value = await pyodide.runPythonAsync(code);
    const durationMs = Math.round(performance.now() - startedAt);
    const result = value === undefined || value === null ? '' : String(value);

    if (value && typeof value === 'object' && 'destroy' in value && typeof value.destroy === 'function') {
      value.destroy();
    }

    self.postMessage({
      type: 'result',
      id,
      result: {
        stdout: stdout.join('\n'),
        stderr: stderr.join('\n'),
        result,
        durationMs,
      } satisfies RunResult,
    });
  } catch (error) {
    self.postMessage({
      type: 'result',
      id,
      result: {
        stdout: stdout.join('\n'),
        stderr: stderr.join('\n'),
        result: '',
        durationMs: 0,
        error: error instanceof Error ? error.message : String(error),
      } satisfies RunResult,
    });
  }
}
