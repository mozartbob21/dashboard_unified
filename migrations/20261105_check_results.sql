-- Extension of run_history. A missing extension is an original (version zero)
-- run with no captured indicators. Never populate it from today's snapshot.
CREATE TABLE IF NOT EXISTS check_results (
 run_id TEXT PRIMARY KEY REFERENCES run_history(run_id) ON DELETE CASCADE,
 version INTEGER NOT NULL, scope TEXT NOT NULL, private_subject TEXT, params TEXT NOT NULL,
 before_signature TEXT, metrics TEXT NOT NULL DEFAULT '{}', result_code TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_run_history_module_time ON run_history(module_id,started_at DESC);
