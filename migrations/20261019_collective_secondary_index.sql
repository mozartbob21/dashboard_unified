-- Forward repair for an intermediate installation with the workspace marker
-- but without this derived lookup table. Apply through store.initialize():
-- it backfills current/historical secondary IDs and records the migration
-- marker in the same transaction. Original records and workflows are untouched.
CREATE TABLE IF NOT EXISTS collective_secondary_ids (
    secondary_id TEXT NOT NULL,
    complaint_key TEXT NOT NULL REFERENCES collective_complaints(key),
    PRIMARY KEY(secondary_id,complaint_key)
);
