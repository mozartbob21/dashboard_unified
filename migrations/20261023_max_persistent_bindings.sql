-- Applies only to data/max_bot.sqlite.
-- 253402300799.0 matches store.PERMANENT_BINDING_EXPIRY: a permanent link,
-- represented by a far-future timestamp so old workers do not delete it while
-- the web application and bot are restarted separately after an upgrade.
-- Preserve routing identity, password fingerprint and the original binding time:
-- presentation subscriptions use created_at to detect an explicit reconnection.
-- Do not revive expired bindings, and never re-run this upgrade after completion.
BEGIN IMMEDIATE;
UPDATE max_bot_bindings
SET expires_at = 253402300799.0
WHERE NOT EXISTS (
    SELECT 1 FROM max_bot_meta WHERE key = 'migration_20261023_max_persistent_bindings'
)
AND typeof(expires_at) IN ('integer', 'real')
AND expires_at > (julianday('now') - 2440587.5) * 86400.0
AND expires_at < 253402300800;
INSERT OR IGNORE INTO max_bot_meta(key,value)
VALUES ('migration_20261023_max_persistent_bindings', '1');
COMMIT;
