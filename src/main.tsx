import ReactDOM from 'react-dom/client';
import App from './App';

// Deliberately not wrapped in StrictMode: it double-invokes effects in dev,
// which would create and tear down a second Univer instance on every mount.
ReactDOM.createRoot(document.getElementById('root') as HTMLElement).render(<App />);
