import { Link } from 'wouter';

export default function NotFound() {
  return (
    <main style={{ minHeight: '100vh', display: 'grid', placeItems: 'center', padding: 24 }}>
      <div style={{ textAlign: 'center' }}>
        <h1>Page not found</h1>
        <Link href="/">Return to Excel Compare</Link>
      </div>
    </main>
  );
}
