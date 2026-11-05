-- Files share the job and dialog lifecycle. Uploaded source books are not retained.
CREATE TABLE IF NOT EXISTS aichat_agent_artifacts (
    id TEXT PRIMARY KEY REFERENCES aichat_jobs(id) ON DELETE CASCADE,
    dialog_id TEXT NOT NULL REFERENCES aichat_dialogs(id) ON DELETE CASCADE,
    contents BLOB NOT NULL CHECK (length(contents) BETWEEN 1 AND 16777216)
);
