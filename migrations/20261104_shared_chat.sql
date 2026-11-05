-- Keep request receipts and confirmations independent of the requesting account.
-- The Python migration adds required_modules to existing aichat_jobs tables.
CREATE TABLE IF NOT EXISTS aichat_request_receipts (
 owner TEXT NOT NULL, request_key TEXT NOT NULL, fingerprint TEXT NOT NULL,
 job_id TEXT NOT NULL,
 PRIMARY KEY(owner,request_key)
);
CREATE TABLE IF NOT EXISTS aichat_confirmations (
 proposal_id TEXT PRIMARY KEY,
 job_id TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS aichat_jobs_module_active ON aichat_jobs(module_id,status);
