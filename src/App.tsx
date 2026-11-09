function App() {
  return (
    <main className="app-shell">
      <section className="welcome-card" aria-labelledby="app-title">
        <p className="eyebrow">BUSINESS MANAGEMENT WORKSPACE</p>
        <h1 id="app-title">Morvarid</h1>
        <p className="welcome-copy">
          Your workspace for import batches, containers, sales, expenses, and
          business reporting.
        </p>
        <div className="status-row">
          <span className="status-dot" aria-hidden="true" />
          Project foundation is ready
        </div>
        <p className="next-step">
          The application foundation is in place. Business features will be
          added in verified, manageable stages.
        </p>
      </section>
    </main>
  );
}

export default App;
