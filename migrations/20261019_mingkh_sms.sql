CREATE TABLE IF NOT EXISTS mingkh_sms_jobs (
    id TEXT PRIMARY KEY,
    owner TEXT NOT NULL,
    action TEXT NOT NULL,
    request_key TEXT NOT NULL,
    fingerprint TEXT NOT NULL,
    payload TEXT NOT NULL,
    status TEXT NOT NULL CHECK(status IN ('queued','running','success','error')),
    error TEXT NOT NULL DEFAULT '',
    created_at REAL NOT NULL,
    updated_at REAL NOT NULL,
    UNIQUE(owner, request_key)
);
CREATE INDEX IF NOT EXISTS mingkh_sms_jobs_queue ON mingkh_sms_jobs(status,created_at);
