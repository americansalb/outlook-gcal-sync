"""SQLite-based sync state persistence."""

import logging
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

from src.sync.models import SyncPair

logger = logging.getLogger("outlook_gcal_sync.sync.state")

SCHEMA_VERSION = 1


class SyncStateStore:
    """Manages the sync state database."""

    def __init__(self, db_path: str):
        self.db_path = db_path
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA journal_mode=WAL")
        self._initialize_schema()

    def _initialize_schema(self) -> None:
        """Create tables if they don't exist."""
        with self.conn:
            self.conn.executescript("""
                CREATE TABLE IF NOT EXISTS sync_pairs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    outlook_id TEXT UNIQUE,
                    google_id TEXT UNIQUE,
                    last_synced_hash TEXT NOT NULL,
                    last_sync_time TEXT NOT NULL,
                    sync_direction TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'synced',
                    outlook_title TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS sync_metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS sync_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT NOT NULL,
                    action TEXT NOT NULL,
                    direction TEXT NOT NULL,
                    outlook_id TEXT,
                    google_id TEXT,
                    event_title TEXT,
                    details TEXT
                );

                CREATE INDEX IF NOT EXISTS idx_sync_pairs_outlook_id
                    ON sync_pairs(outlook_id);
                CREATE INDEX IF NOT EXISTS idx_sync_pairs_google_id
                    ON sync_pairs(google_id);
            """)

            # Store schema version
            self.conn.execute(
                "INSERT OR REPLACE INTO sync_metadata (key, value) VALUES (?, ?)",
                ("schema_version", str(SCHEMA_VERSION)),
            )

    def close(self) -> None:
        """Close the database connection."""
        self.conn.close()

    def get_pair_by_outlook_id(self, outlook_id: str) -> SyncPair | None:
        """Find a sync pair by Outlook event ID."""
        row = self.conn.execute(
            "SELECT * FROM sync_pairs WHERE outlook_id = ?", (outlook_id,),
        ).fetchone()
        return self._row_to_pair(row) if row else None

    def get_pair_by_google_id(self, google_id: str) -> SyncPair | None:
        """Find a sync pair by Google event ID."""
        row = self.conn.execute(
            "SELECT * FROM sync_pairs WHERE google_id = ?", (google_id,),
        ).fetchone()
        return self._row_to_pair(row) if row else None

    def get_all_pairs(self) -> list[SyncPair]:
        """Get all sync pairs."""
        rows = self.conn.execute("SELECT * FROM sync_pairs").fetchall()
        return [self._row_to_pair(row) for row in rows]

    def upsert_pair(self, pair: SyncPair) -> int:
        """Insert or update a sync pair. Returns the row ID."""
        now = datetime.now(timezone.utc).isoformat()
        with self.conn:
            if pair.id is not None:
                self.conn.execute(
                    """UPDATE sync_pairs
                       SET outlook_id = ?, google_id = ?, last_synced_hash = ?,
                           last_sync_time = ?, sync_direction = ?, status = ?,
                           outlook_title = ?, updated_at = ?
                       WHERE id = ?""",
                    (pair.outlook_id, pair.google_id, pair.last_synced_hash,
                     pair.last_sync_time, pair.sync_direction, pair.status,
                     pair.outlook_title, now, pair.id),
                )
                return pair.id
            else:
                cursor = self.conn.execute(
                    """INSERT INTO sync_pairs
                       (outlook_id, google_id, last_synced_hash, last_sync_time,
                        sync_direction, status, outlook_title, created_at, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (pair.outlook_id, pair.google_id, pair.last_synced_hash,
                     pair.last_sync_time, pair.sync_direction, pair.status,
                     pair.outlook_title, now, now),
                )
                return cursor.lastrowid  # type: ignore[return-value]

    def delete_pair(self, pair_id: int) -> None:
        """Delete a sync pair by ID."""
        with self.conn:
            self.conn.execute("DELETE FROM sync_pairs WHERE id = ?", (pair_id,))

    def delete_pair_by_outlook_id(self, outlook_id: str) -> None:
        """Delete a sync pair by Outlook event ID."""
        with self.conn:
            self.conn.execute("DELETE FROM sync_pairs WHERE outlook_id = ?", (outlook_id,))

    def delete_pair_by_google_id(self, google_id: str) -> None:
        """Delete a sync pair by Google event ID."""
        with self.conn:
            self.conn.execute("DELETE FROM sync_pairs WHERE google_id = ?", (google_id,))

    def get_metadata(self, key: str) -> str | None:
        """Get a metadata value."""
        row = self.conn.execute(
            "SELECT value FROM sync_metadata WHERE key = ?", (key,),
        ).fetchone()
        return row["value"] if row else None

    def set_metadata(self, key: str, value: str) -> None:
        """Set a metadata value."""
        with self.conn:
            self.conn.execute(
                "INSERT OR REPLACE INTO sync_metadata (key, value) VALUES (?, ?)",
                (key, value),
            )

    def log_action(
        self,
        action: str,
        direction: str,
        outlook_id: str | None = None,
        google_id: str | None = None,
        event_title: str = "",
        details: str = "",
    ) -> None:
        """Log a sync action for auditing."""
        now = datetime.now(timezone.utc).isoformat()
        with self.conn:
            self.conn.execute(
                """INSERT INTO sync_log
                   (timestamp, action, direction, outlook_id, google_id, event_title, details)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (now, action, direction, outlook_id, google_id, event_title, details),
            )

    def get_stats(self) -> dict:
        """Get sync statistics for the status command."""
        total = self.conn.execute("SELECT COUNT(*) FROM sync_pairs").fetchone()[0]
        last_sync = self.get_metadata("last_sync_time")
        recent_actions = self.conn.execute(
            "SELECT action, COUNT(*) as cnt FROM sync_log "
            "WHERE timestamp > datetime('now', '-1 day') GROUP BY action"
        ).fetchall()
        return {
            "total_pairs": total,
            "last_sync_time": last_sync,
            "recent_actions": {row["action"]: row["cnt"] for row in recent_actions},
        }

    def reset(self) -> None:
        """Clear all sync state. Used by the reset command."""
        with self.conn:
            self.conn.execute("DELETE FROM sync_pairs")
            self.conn.execute("DELETE FROM sync_metadata WHERE key != 'schema_version'")
            self.conn.execute("DELETE FROM sync_log")
        logger.warning("Sync state has been reset.")

    @staticmethod
    def _row_to_pair(row: sqlite3.Row) -> SyncPair:
        """Convert a database row to a SyncPair."""
        return SyncPair(
            id=row["id"],
            outlook_id=row["outlook_id"],
            google_id=row["google_id"],
            last_synced_hash=row["last_synced_hash"],
            last_sync_time=row["last_sync_time"],
            sync_direction=row["sync_direction"],
            status=row["status"],
            outlook_title=row["outlook_title"] or "",
        )
