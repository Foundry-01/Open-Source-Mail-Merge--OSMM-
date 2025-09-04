import React from 'react';
import { Routes, Route, Link } from 'react-router-dom';
import Home from './Home';
import Privacy from './Privacy';
import Terms from './Terms';

export default function App(): JSX.Element {
  return (
    <div className="site">
      <header className="site-header">
        <div className="container">
          <Link to="/" className="brand">
            <img src="/logo.svg" alt="OSMM" className="brand-logo" />
            <span className="brand-name">Open Source Mail Merge</span>
            <span className="brand-name-short">OSMM</span>
          </Link>
          <nav className="nav">
            <Link to="/" className="nav-link">Home</Link>
            <Link to="/privacy" className="nav-link">Privacy</Link>
            <a href="https://github.com/jdragon3001/Open-Source-Mail-Merge--OSMM-" className="nav-link" target="_blank" rel="noreferrer">Github</a>
          </nav>
        </div>
      </header>
      <main className="container">
        <Routes>
          <Route path="/" element={<Home />} />
          <Route path="/privacy" element={<Privacy />} />
          <Route path="/terms" element={<Terms />} />
        </Routes>
      </main>
      <footer className="site-footer">
        <div className="container small">
          <div>© {new Date().getFullYear()} OSMM • Open-source under Apache-2.0</div>
          <div className="footer-links">
            <Link to="/privacy" className="link">Privacy Policy</Link>
            <span>•</span>
            <Link to="/terms" className="link">Terms of Service</Link>
          </div>
        </div>
      </footer>
    </div>
  );
}


