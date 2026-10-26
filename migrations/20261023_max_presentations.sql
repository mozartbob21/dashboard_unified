-- Private MAX bot DB only. No passwords, document bodies or upload tokens.
CREATE TABLE IF NOT EXISTS max_presentations_state (
    id INTEGER PRIMARY KEY CHECK (id=1), baseline_at REAL, last_checked REAL,
    last_new REAL, next_scan REAL NOT NULL DEFAULT 0, last_error TEXT,
    failures INTEGER NOT NULL DEFAULT 0
);
INSERT OR IGNORE INTO max_presentations_state(id) VALUES(1);
CREATE TABLE IF NOT EXISTS max_presentations_files (
    attachment_id TEXT PRIMARY KEY, name TEXT NOT NULL, description TEXT NOT NULL,
    created_at TEXT NOT NULL, source_path TEXT NOT NULL, discovered_at REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS max_presentations_subscriptions (
    user_id INTEGER PRIMARY KEY, max_user_id TEXT NOT NULL, chat_id TEXT NOT NULL,
    binding_created_at REAL NOT NULL, generation INTEGER NOT NULL,
    enabled INTEGER NOT NULL CHECK(enabled IN (0,1)), activated_at REAL NOT NULL,
    updated_at REAL NOT NULL, last_error TEXT
);
CREATE TABLE IF NOT EXISTS max_presentations_deliveries (
    id INTEGER PRIMARY KEY, attachment_id TEXT NOT NULL, user_id INTEGER NOT NULL,
    generation INTEGER NOT NULL, state TEXT NOT NULL CHECK(state IN
      ('pending','uploading','sending','retry','sent','uncertain','failed','cancelled')),
    attempts INTEGER NOT NULL DEFAULT 0, next_attempt REAL NOT NULL DEFAULT 0,
    last_error TEXT, created_at REAL NOT NULL, updated_at REAL NOT NULL,
    UNIQUE(attachment_id,user_id,generation)
);
CREATE INDEX IF NOT EXISTS max_presentations_due ON max_presentations_deliveries(state,next_attempt,id);
CREATE TABLE IF NOT EXISTS max_presentations_lease (
    id INTEGER PRIMARY KEY CHECK(id=1), owner TEXT NOT NULL, expires_at REAL NOT NULL
);
