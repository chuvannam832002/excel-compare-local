import { createRoot } from 'react-dom/client';

import App from './App';
import PythonRunner from '@/pages/PythonRunner';
import NotFound from '@/pages/not-found';
import { ErrorBoundary } from '@/components/error-boundary';
import { Route, Router as WouterRouter, Switch } from 'wouter';

import './index.css';

function AppRouter() {
  return (
    <WouterRouter base={import.meta.env.BASE_URL.replace(/\/$/, '')}>
      <Switch>
        <Route path="/" component={App} />
        <Route path="/python" component={PythonRunner} />
        <Route component={NotFound} />
      </Switch>
    </WouterRouter>
  );
}

createRoot(document.getElementById('root')!, {
  // Keeps caught errors off reportError(), which would raise the dev overlay.
  onCaughtError: (error, errorInfo) => {
    console.error(error, errorInfo.componentStack);
  },
}).render(
  <ErrorBoundary>
    <AppRouter />
  </ErrorBoundary>,
);
