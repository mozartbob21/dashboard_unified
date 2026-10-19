CREATE TABLE IF NOT EXISTS collective_refresh_jobs (
    id TEXT PRIMARY KEY,
    start TEXT NOT NULL,
    end TEXT NOT NULL,
    status TEXT NOT NULL CHECK(status IN ('queued','running','success','error')),
    error TEXT NOT NULL DEFAULT '',
    created_at REAL NOT NULL,
    updated_at REAL NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS collective_one_refresh
    ON collective_refresh_jobs((1)) WHERE status IN ('queued','running');

CREATE TABLE IF NOT EXISTS collective_refresh_requests (
    request_key TEXT PRIMARY KEY,
    job_id TEXT NOT NULL REFERENCES collective_refresh_jobs(id) ON DELETE CASCADE
);
