-- Applies only to data/max_bot.sqlite, NEVER to dashboard.db.
BEGIN IMMEDIATE;
CREATE TABLE IF NOT EXISTS max_bot_meta(key TEXT PRIMARY KEY, value TEXT NOT NULL);
INSERT OR IGNORE INTO max_bot_meta(key,value) VALUES ('schema_version','1');
CREATE TABLE IF NOT EXISTS max_bot_codes (
 code_hash TEXT PRIMARY KEY, user_id INTEGER NOT NULL, password_fingerprint TEXT NOT NULL,
 created_at REAL NOT NULL, expires_at REAL NOT NULL, consumed INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS max_bot_code_owner ON max_bot_codes(user_id);
CREATE TABLE IF NOT EXISTS max_bot_pending (
 request_id TEXT PRIMARY KEY, user_id INTEGER NOT NULL, max_user_id TEXT NOT NULL,
 chat_id TEXT NOT NULL, password_fingerprint TEXT NOT NULL, created_at REAL NOT NULL, expires_at REAL NOT NULL
);
CREATE INDEX IF NOT EXISTS max_bot_pending_owner ON max_bot_pending(user_id);
CREATE TABLE IF NOT EXISTS max_bot_bindings (
 user_id INTEGER PRIMARY KEY, max_user_id TEXT NOT NULL UNIQUE, chat_id TEXT NOT NULL,
 password_fingerprint TEXT NOT NULL, created_at REAL NOT NULL, expires_at REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS max_bot_rate_events (
 id INTEGER PRIMARY KEY, bucket TEXT NOT NULL, key_hash TEXT NOT NULL, occurred_at REAL NOT NULL
);
CREATE INDEX IF NOT EXISTS max_bot_rate_window ON max_bot_rate_events(bucket,key_hash,occurred_at);
CREATE TABLE IF NOT EXISTS max_bot_audit (
 id INTEGER PRIMARY KEY, occurred_at REAL NOT NULL, event TEXT NOT NULL, outcome TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS max_bot_lease (
 id INTEGER PRIMARY KEY CHECK(id=1), owner TEXT NOT NULL, expires_at REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS max_bot_events (
 event_hash TEXT PRIMARY KEY, status TEXT NOT NULL, created_at REAL NOT NULL, finished_at REAL
);
COMMIT;
