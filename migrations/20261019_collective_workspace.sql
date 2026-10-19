-- Shared, authenticated Collective workspace. No source credentials or browser-local data.
CREATE TABLE IF NOT EXISTS collective_settings (
    id INTEGER PRIMARY KEY CHECK(id=1),
    since TEXT NOT NULL,
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS collective_assignees (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    name_key TEXT NOT NULL UNIQUE,
    archived INTEGER NOT NULL DEFAULT 0 CHECK(archived IN (0,1)),
    version INTEGER NOT NULL DEFAULT 1 CHECK(version>=1),
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    updated_by TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS collective_complaints (
    key TEXT PRIMARY KEY,
    source_json TEXT NOT NULL,
    first_seen_at TEXT NOT NULL,
    last_seen_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS collective_identity_aliases (
    alias TEXT PRIMARY KEY,
    complaint_key TEXT NOT NULL REFERENCES collective_complaints(key)
);
CREATE INDEX IF NOT EXISTS collective_aliases_complaint ON collective_identity_aliases(complaint_key);
CREATE TABLE IF NOT EXISTS collective_secondary_ids (
    secondary_id TEXT NOT NULL,
    complaint_key TEXT NOT NULL REFERENCES collective_complaints(key),
    PRIMARY KEY(secondary_id,complaint_key)
);
CREATE TABLE IF NOT EXISTS collective_workflows (
    complaint_key TEXT PRIMARY KEY REFERENCES collective_complaints(key),
    version INTEGER NOT NULL DEFAULT 0 CHECK(version>=0),
    assignee_id INTEGER REFERENCES collective_assignees(id),
    report TEXT NOT NULL DEFAULT '' CHECK(length(report)<=20000),
    theme TEXT CHECK(theme IN ('ii','kr','pr') OR theme IS NULL),
    assigned_at TEXT,
    assigned_by TEXT NOT NULL DEFAULT '',
    report_at TEXT,
    report_by TEXT NOT NULL DEFAULT '',
    theme_at TEXT,
    theme_by TEXT NOT NULL DEFAULT '',
    updated_at TEXT,
    updated_by TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS collective_workflow_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    complaint_key TEXT NOT NULL REFERENCES collective_complaints(key),
    version INTEGER NOT NULL,
    body_json TEXT NOT NULL,
    recorded_at TEXT NOT NULL,
    recorded_by TEXT NOT NULL DEFAULT '',
    UNIQUE(complaint_key,version)
);
CREATE INDEX IF NOT EXISTS collective_history_complaint ON collective_workflow_history(complaint_key,version);
CREATE TABLE IF NOT EXISTS collective_snapshots (
    start TEXT NOT NULL,
    end TEXT NOT NULL,
    rows_json TEXT NOT NULL,
    keys_json TEXT NOT NULL,
    kurators_json TEXT NOT NULL,
    meta TEXT NOT NULL,
    fetched_at TEXT NOT NULL,
    version INTEGER NOT NULL DEFAULT 1 CHECK(version>=1),
    PRIMARY KEY(start,end)
);
CREATE INDEX IF NOT EXISTS collective_snapshots_fetched ON collective_snapshots(fetched_at);
CREATE TABLE IF NOT EXISTS app_schema_migrations (
    name TEXT PRIMARY KEY,
    applied_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
INSERT OR IGNORE INTO app_schema_migrations(name) VALUES ('20261019_collective_workspace');
