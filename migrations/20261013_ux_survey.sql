-- Applied automatically by utils.db; safe to re-run on an existing database.
BEGIN IMMEDIATE;
CREATE TABLE IF NOT EXISTS app_schema_migrations (
    name TEXT PRIMARY KEY,
    applied_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS ux_survey_state (
    account_key TEXT NOT NULL,
    survey_id TEXT NOT NULL,
    first_shown_at REAL,
    last_shown_at REAL,
    next_invite_at REAL,
    invitation_token TEXT,
    draft_q1 INTEGER CHECK(draft_q1 IS NULL OR (typeof(draft_q1)='integer' AND draft_q1 BETWEEN 0 AND 5)),
    draft_q2 INTEGER CHECK(draft_q2 IS NULL OR (typeof(draft_q2)='integer' AND draft_q2 BETWEEN 0 AND 5)),
    draft_q3 INTEGER CHECK(draft_q3 IS NULL OR (typeof(draft_q3)='integer' AND draft_q3 BETWEEN 0 AND 5)),
    draft_q4 INTEGER CHECK(draft_q4 IS NULL OR (typeof(draft_q4)='integer' AND draft_q4 BETWEEN 0 AND 5)),
    draft_q5 INTEGER CHECK(draft_q5 IS NULL OR (typeof(draft_q5)='integer' AND draft_q5 BETWEEN 0 AND 5)),
    draft_q6 INTEGER CHECK(draft_q6 IS NULL OR (typeof(draft_q6)='integer' AND draft_q6 BETWEEN 0 AND 5)),
    draft_comment TEXT NOT NULL DEFAULT '' CHECK(length(draft_comment)<=2000),
    draft_version INTEGER NOT NULL DEFAULT 0,
    draft_updated_at REAL,
    PRIMARY KEY(account_key, survey_id)
);
CREATE TABLE IF NOT EXISTS ux_survey_responses (
    account_key TEXT NOT NULL,
    survey_id TEXT NOT NULL,
    q1 INTEGER NOT NULL CHECK(typeof(q1)='integer' AND q1 BETWEEN 0 AND 5),
    q2 INTEGER NOT NULL CHECK(typeof(q2)='integer' AND q2 BETWEEN 0 AND 5),
    q3 INTEGER NOT NULL CHECK(typeof(q3)='integer' AND q3 BETWEEN 0 AND 5),
    q4 INTEGER NOT NULL CHECK(typeof(q4)='integer' AND q4 BETWEEN 0 AND 5),
    q5 INTEGER NOT NULL CHECK(typeof(q5)='integer' AND q5 BETWEEN 0 AND 5),
    q6 INTEGER NOT NULL CHECK(typeof(q6)='integer' AND q6 BETWEEN 0 AND 5),
    comment TEXT NOT NULL DEFAULT '' CHECK(length(comment)<=2000),
    submitted_at REAL NOT NULL,
    PRIMARY KEY(account_key, survey_id),
    FOREIGN KEY(account_key, survey_id) REFERENCES ux_survey_state(account_key, survey_id)
);
CREATE INDEX IF NOT EXISTS ux_survey_responses_date ON ux_survey_responses(survey_id, submitted_at);
CREATE TRIGGER IF NOT EXISTS ux_survey_responses_immutable
BEFORE UPDATE ON ux_survey_responses
BEGIN
    SELECT RAISE(ABORT, 'Submitted survey responses are immutable');
END;
INSERT OR IGNORE INTO app_schema_migrations(name) VALUES ('20261013_ux_survey');
COMMIT;
