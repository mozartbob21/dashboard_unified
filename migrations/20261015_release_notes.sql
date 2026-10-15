-- Applied lazily by services.release_notes, or explicitly with initialize().
-- Independent from account_notifications: these receipts belong to each account.
BEGIN IMMEDIATE;
CREATE TABLE IF NOT EXISTS app_schema_migrations (
    name TEXT PRIMARY KEY,
    applied_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS release_note_receipts (
    account_key TEXT NOT NULL CHECK(length(account_key) BETWEEN 1 AND 300),
    release_id TEXT NOT NULL CHECK(length(release_id) BETWEEN 1 AND 100),
    read_at TEXT NOT NULL,
    PRIMARY KEY(account_key, release_id)
);
INSERT OR IGNORE INTO app_schema_migrations(name) VALUES ('20261015_release_notes');
COMMIT;
