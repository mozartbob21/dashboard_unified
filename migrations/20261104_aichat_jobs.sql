-- Existing shared JSON is preserved, but cannot be assigned to an account safely.
CREATE TABLE IF NOT EXISTS aichat_dialogs (
 id TEXT PRIMARY KEY, owner TEXT NOT NULL, body TEXT NOT NULL, updated_at REAL NOT NULL
);
CREATE INDEX IF NOT EXISTS aichat_dialog_owner ON aichat_dialogs(owner,updated_at);
CREATE TABLE IF NOT EXISTS aichat_jobs (
 id TEXT PRIMARY KEY, owner TEXT NOT NULL, request_key TEXT NOT NULL,
 fingerprint TEXT NOT NULL, dialog_id TEXT NOT NULL,
 status TEXT NOT NULL, stage TEXT NOT NULL, steps TEXT NOT NULL DEFAULT '[]',
 result TEXT, module_id TEXT, created_at REAL NOT NULL, updated_at REAL NOT NULL,
 heartbeat REAL NOT NULL, deadline REAL NOT NULL,
 UNIQUE(owner,request_key)
);
CREATE INDEX IF NOT EXISTS aichat_jobs_owner ON aichat_jobs(owner,dialog_id,created_at);
