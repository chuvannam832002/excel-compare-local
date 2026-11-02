export type PythonRunStatus = 'loading' | 'running';

export type PythonRunResult = {
  stdout: string;
  stderr: string;
  result: string;
  durationMs: number;
  error?: string;
};

type WorkerStatusMessage = {
  type: 'status';
  id: number;
  status: PythonRunStatus;
};

type WorkerResultMessage = {
  type: 'result';
  id: number;
  result: PythonRunResult;
};

type PendingRun = {
  id: number;
  resolve: (result: PythonRunResult) => void;
  reject: (error: Error) => void;
  onStatus?: (status: PythonRunStatus) => void;
  timer: ReturnType<typeof setTimeout>;
};

let runtimeWorker: Worker | null = null;
let pendingRun: PendingRun | null = null;
let nextRunId = 0;

function clearPendingTimer(run: PendingRun) {
  clearTimeout(run.timer);
}

function stopWorker() {
  runtimeWorker?.terminate();
  runtimeWorker = null;
}

function handleTimeout(id: number, message: string) {
  if (!pendingRun || pendingRun.id !== id) return;
  const run = pendingRun;
  pendingRun = null;
  clearPendingTimer(run);
  stopWorker();
  run.resolve({
    stdout: '',
    stderr: '',
    result: '',
    durationMs: message.includes('15 seconds') ? 15000 : 0,
    error: message,
  });
}

function getWorker() {
  if (runtimeWorker) return runtimeWorker;

  const worker = new Worker(new URL('../workers/python.worker.ts', import.meta.url), {
    type: 'module',
  });

  worker.onmessage = (event: MessageEvent<WorkerStatusMessage | WorkerResultMessage>) => {
    const message = event.data;
    if (!pendingRun || pendingRun.id !== message.id) return;

    if (message.type === 'status') {
      pendingRun.onStatus?.(message.status);
      clearPendingTimer(pendingRun);
      const currentRun = pendingRun;
      const timeoutMs = message.status === 'loading' ? 120000 : 15000;
      currentRun.timer = setTimeout(
        () => handleTimeout(currentRun.id, message.status === 'loading'
          ? 'Python took too long to load. Check your internet connection and try again.'
          : 'Python execution timed out after 15 seconds. The worker was stopped.'),
        timeoutMs,
      );
      return;
    }

    const run = pendingRun;
    pendingRun = null;
    clearPendingTimer(run);
    run.resolve(message.result);
  };

  worker.onerror = (event) => {
    if (pendingRun) {
      const run = pendingRun;
      pendingRun = null;
      clearPendingTimer(run);
      run.reject(new Error(event.message || 'The Python worker could not start.'));
    }
    stopWorker();
  };

  runtimeWorker = worker;
  return worker;
}

export function runPython(
  code: string,
  onStatus?: (status: PythonRunStatus) => void,
): Promise<PythonRunResult> {
  if (pendingRun) {
    return Promise.reject(new Error('A Python script is already running.'));
  }

  return new Promise((resolve, reject) => {
    const id = ++nextRunId;
    const timer = setTimeout(
      () => handleTimeout(id, 'Python took too long to initialize. Please try again.'),
      120000,
    );
    pendingRun = { id, resolve, reject, onStatus, timer };
    getWorker().postMessage({ type: 'run', id, code });
  });
}
