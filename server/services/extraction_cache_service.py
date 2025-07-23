"""
Extraction Cache Service for optimizing file processing performance.
Stores extracted data temporarily to avoid redundant PDF/Excel parsing.
"""

import uuid
import time
import gc
import logging
from typing import Dict, List, Optional, Any
from threading import Lock
from ..schemas.tender import TenderItem, SubtableItem

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ExtractionCacheService:
    """
    Thread-safe cache service for storing extracted PDF and Excel data.
    Uses in-memory storage with automatic cleanup after expiration.
    """

    def __init__(self, default_ttl_minutes: int = 30):
        """
        Initialize the cache service.

        Args:
            default_ttl_minutes: Default time-to-live for cached data in minutes
        """
        self._cache: Dict[str, Dict[str, Any]] = {}
        self._lock = Lock()
        self.default_ttl = default_ttl_minutes * 60  # Convert to seconds

    def store_extraction_results(
        self,
        pdf_items: List[TenderItem],
        excel_items: List[TenderItem],
        pdf_subtables: List[SubtableItem],
        excel_subtables: List[SubtableItem],
        extraction_params: Dict[str, Any],
        session_id: Optional[str] = None
    ) -> str:
        """
        Store extraction results with a unique session ID.

        Args:
            pdf_items: Extracted PDF main table items
            excel_items: Extracted Excel main table items
            pdf_subtables: Extracted PDF subtable items
            excel_subtables: Extracted Excel subtable items
            extraction_params: Parameters used for extraction
            session_id: Optional custom session ID

        Returns:
            session_id: Unique identifier for this extraction session
        """
        if session_id is None:
            session_id = str(uuid.uuid4())

        current_time = time.time()

        with self._lock:
            self._cache[session_id] = {
                'pdf_items': pdf_items,
                'excel_items': excel_items,
                'pdf_subtables': pdf_subtables,
                'excel_subtables': excel_subtables,
                'extraction_params': extraction_params,
                'created_at': current_time,
                'expires_at': current_time + self.default_ttl,
                'access_count': 0,
                'last_accessed': current_time
            }

        logger.info(f"Stored extraction results for session {session_id}")
        logger.info(f"Cache contains: {len(pdf_items)} PDF items, {len(excel_items)} Excel items, "
                    f"{len(pdf_subtables)} PDF subtables, {len(excel_subtables)} Excel subtables")

        return session_id

    def get_extraction_results(self, session_id: str) -> Optional[Dict[str, Any]]:
        """
        Retrieve extraction results by session ID.

        Args:
            session_id: Session identifier

        Returns:
            Dictionary containing extraction results or None if not found/expired
        """
        current_time = time.time()

        with self._lock:
            if session_id not in self._cache:
                logger.warning(
                    f"No extraction data found for session {session_id}")
                return None

            cache_entry = self._cache[session_id]

            # Check if expired
            if current_time > cache_entry['expires_at']:
                logger.warning(
                    f"Extraction data for session {session_id} has expired")
                del self._cache[session_id]
                return None

            # Update access tracking
            cache_entry['access_count'] += 1
            cache_entry['last_accessed'] = current_time

            logger.info(f"Retrieved extraction results for session {session_id} "
                        f"(access count: {cache_entry['access_count']})")

            return {
                'pdf_items': cache_entry['pdf_items'],
                'excel_items': cache_entry['excel_items'],
                'pdf_subtables': cache_entry['pdf_subtables'],
                'excel_subtables': cache_entry['excel_subtables'],
                'extraction_params': cache_entry['extraction_params']
            }

    def extend_session(self, session_id: str, additional_minutes: int = 30) -> bool:
        """
        Extend the expiration time for a session.

        Args:
            session_id: Session identifier
            additional_minutes: Additional time to add in minutes

        Returns:
            True if session was extended, False if session not found
        """
        current_time = time.time()
        additional_seconds = additional_minutes * 60

        with self._lock:
            if session_id not in self._cache:
                return False

            cache_entry = self._cache[session_id]
            cache_entry['expires_at'] = max(
                cache_entry['expires_at'],
                current_time + additional_seconds
            )

            logger.info(
                f"Extended session {session_id} by {additional_minutes} minutes")
            return True

    def cleanup_session(self, session_id: str) -> bool:
        """
        Manually cleanup a specific session.

        Args:
            session_id: Session identifier

        Returns:
            True if session was cleaned up, False if not found
        """
        with self._lock:
            if session_id in self._cache:
                del self._cache[session_id]
                logger.info(f"Cleaned up session {session_id}")
                gc.collect()  # Force garbage collection
                return True
            return False

    def cleanup_expired_sessions(self) -> int:
        """
        Remove all expired sessions from cache.

        Returns:
            Number of sessions cleaned up
        """
        current_time = time.time()
        expired_sessions = []

        with self._lock:
            for session_id, cache_entry in self._cache.items():
                if current_time > cache_entry['expires_at']:
                    expired_sessions.append(session_id)

            for session_id in expired_sessions:
                del self._cache[session_id]

        if expired_sessions:
            logger.info(f"Cleaned up {len(expired_sessions)} expired sessions")
            gc.collect()

        return len(expired_sessions)

    def get_cache_stats(self) -> Dict[str, Any]:
        """
        Get cache statistics.

        Returns:
            Dictionary with cache statistics
        """
        current_time = time.time()

        with self._lock:
            active_sessions = 0
            expired_sessions = 0
            total_items = 0
            total_subtables = 0

            for cache_entry in self._cache.values():
                if current_time <= cache_entry['expires_at']:
                    active_sessions += 1
                    total_items += len(cache_entry['pdf_items']) + \
                        len(cache_entry['excel_items'])
                    total_subtables += len(cache_entry['pdf_subtables']) + len(
                        cache_entry['excel_subtables'])
                else:
                    expired_sessions += 1

            return {
                'active_sessions': active_sessions,
                'expired_sessions': expired_sessions,
                'total_sessions': len(self._cache),
                'total_items_cached': total_items,
                'total_subtables_cached': total_subtables
            }


# Global cache instance
_extraction_cache = ExtractionCacheService()


def get_extraction_cache() -> ExtractionCacheService:
    """Get the global extraction cache instance."""
    return _extraction_cache
