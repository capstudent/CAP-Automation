import threading
import time

from backend.sheets_service import SheetsService
from backend.automation_service import AutomationService


class UserSession:
    """Holds one user's sheets + automation services and tracks activity."""

    def __init__(self, config):
        self.sheets_service = SheetsService(config)
        self.automation_service = AutomationService(config, self.sheets_service)
        self.created_at = time.time()
        self.last_activity = time.time()

    def touch(self):
        self.last_activity = time.time()

    def cleanup(self):
        try:
            self.automation_service._cleanup_driver()
        except Exception as e:
            print(f"⚠️  Error cleaning up session: {e}")


class SessionManager:
    """Manages per-user UserSession instances with idle timeout and capacity cap."""

    def __init__(self, config, max_sessions=10, idle_timeout_seconds=1800):
        self.config = config
        self.max_sessions = max_sessions
        self.idle_timeout = idle_timeout_seconds
        self._sessions = {}
        self._lock = threading.Lock()
        self._start_cleanup_thread()

    def get_or_create(self, session_id):
        with self._lock:
            if session_id in self._sessions:
                self._sessions[session_id].touch()
                return self._sessions[session_id]

            # Purge idle sessions before checking capacity
            self._cleanup_idle_locked()

            if len(self._sessions) >= self.max_sessions:
                raise Exception(
                    f"Server is at capacity ({self.max_sessions} concurrent users). "
                    f"Please try again in a few minutes."
                )

            print(f"🆕 Creating new session {session_id[:8]}... ({len(self._sessions) + 1}/{self.max_sessions} active)")
            user_session = UserSession(self.config)
            self._sessions[session_id] = user_session
            return user_session

    def remove(self, session_id):
        with self._lock:
            user_session = self._sessions.pop(session_id, None)
        if user_session:
            print(f"🗑️  Removing session {session_id[:8]}...")
            user_session.cleanup()

    def active_count(self):
        with self._lock:
            return len(self._sessions)

    def _cleanup_idle_locked(self):
        """Must be called with self._lock held."""
        now = time.time()
        stale_ids = [
            sid for sid, s in self._sessions.items()
            if now - s.last_activity > self.idle_timeout
        ]
        for sid in stale_ids:
            print(f"🧹 Session {sid[:8]} idle >{self.idle_timeout}s, cleaning up")
            sess = self._sessions.pop(sid, None)
            if sess:
                try:
                    sess.cleanup()
                except Exception as e:
                    print(f"   Cleanup error: {e}")

    def _start_cleanup_thread(self):
        def loop():
            while True:
                time.sleep(60)
                try:
                    with self._lock:
                        self._cleanup_idle_locked()
                except Exception as e:
                    print(f"Cleanup thread error: {e}")

        thread = threading.Thread(target=loop, daemon=True)
        thread.start()
