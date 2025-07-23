import jaconv
import re
from typing import Dict


class Normalizer:
    def __init__(self):
        # Common synonyms mapping for Japanese construction terms
        self.synonyms = {
            "工事区分": "工種",
            "工種区分": "工種",
            "種別区分": "種別",
            "細目": "細別",
            "名称・規格": "名称",
            "品名": "名称",
            "項目": "名称",
            "数量": "数量"
        }

    def normalize_key(self, fields: Dict[str, str]) -> str:
        """
        Normalize item fields into a single comparable key.
        """
        normalized_parts = []

        for field_name, value in fields.items():
            if not value:
                continue

            # Map synonyms first
            normalized_field_name = self.synonyms.get(field_name, field_name)

            # Normalize the value
            normalized_value = self._normalize_text(value)

            if normalized_value:
                normalized_parts.append(
                    f"{normalized_field_name}:{normalized_value}")

        return "|".join(sorted(normalized_parts))

    def normalize_item(self, item_key: str) -> str:
        """
        Normalize a single item key string.
        This should handle the pipe-separated format created by parsers.
        """
        if not item_key:
            return ""

        # If it's already a structured key (contains |), normalize each part
        if "|" in item_key:
            parts = item_key.split("|")
            normalized_parts = []
            for part in parts:
                normalized_part = self._normalize_text(part)
                if normalized_part:
                    normalized_parts.append(normalized_part)
            return "|".join(normalized_parts)
        else:
            # Single string, just normalize it
            return self._normalize_text(item_key)

    def _normalize_text(self, text: str) -> str:
        """
        Enhanced text normalization function with better handling of technical specifications
        """
        if not text:
            return ""

        # Convert to string if not already
        text = str(text)

        # Convert full-width to half-width
        normalized = jaconv.z2h(text, kana=False, digit=True, ascii=True)

        # Convert to lowercase
        normalized = normalized.lower()

        # Remove row spanning artifacts (+ symbols from concatenation)
        normalized = re.sub(r'\s*\+\s*', '', normalized)

        # Remove various types of spaces and special characters
        # Remove spaces including full-width
        normalized = re.sub(r'[\s　\u3000]+', '', normalized)

        # Keep alphanumeric, Japanese characters, and important technical symbols
        # Preserve × (multiplication), ✕ (cross), φ (phi), * (asterisk), ~ (wave dash), commas, and other technical symbols
        normalized = re.sub(
            r'[^\w\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAFφΦ×✕*～〜,，、。]', '', normalized)

        # Remove common noise words/characters but preserve numerical suffixes
        noise_patterns = [
            r'^第\d+号',  # Remove "第X号"
            r'当り$',     # Remove "当り" at end
            r'当たり$',   # Remove "当たり" at end
        ]

        for pattern in noise_patterns:
            normalized = re.sub(pattern, '', normalized)

        return normalized.strip()

    def are_items_significantly_different(self, text1: str, text2: str) -> bool:
        """
        STRICT EXACT MATCHING: Check if two items are different using exact word matching.
        Items are considered the same ONLY if they match exactly after whitespace normalization.
        This ensures precise identification for construction documents.
        """
        if not text1 or not text2:
            return True

        import re

        def normalize_for_exact_match(text):
            """
            Normalize text for exact matching by:
            1. Converting full-width to half-width
            2. Removing all whitespace
            3. Converting to lowercase
            4. Preserving all meaningful characters
            """
            if not text:
                return ""

            # Convert to string if not already
            text = str(text)

            # Convert full-width to half-width (for numbers and basic ASCII)
            normalized = jaconv.z2h(text, kana=False, digit=True, ascii=True)

            # Remove all types of whitespace (spaces, tabs, full-width spaces, etc.)
            normalized = re.sub(r'[\s　\u3000\t\n\r]+', '', normalized)

            # Convert to lowercase for case-insensitive comparison
            normalized = normalized.lower()

            return normalized.strip()

        # Normalize both texts for exact comparison
        norm1 = normalize_for_exact_match(text1)
        norm2 = normalize_for_exact_match(text2)

        # EXACT MATCH: Items are the same ONLY if they match exactly
        if norm1 == norm2:
            return False  # Not significantly different - they're the same
        else:
            return True   # Significantly different - they're different items

    def calculate_similarity_score(self, key1: str, key2: str) -> float:
        """
        Calculate similarity score between two normalized keys.
        Returns a value between 0.0 and 1.0.
        """
        if not key1 or not key2:
            return 0.0

        norm1 = self.normalize_item(key1)
        norm2 = self.normalize_item(key2)

        if norm1 == norm2:
            return 1.0

        # Simple character-based similarity
        if len(norm1) == 0 or len(norm2) == 0:
            return 0.0

        # Count common characters
        common_chars = sum(1 for c in norm1 if c in norm2)
        max_length = max(len(norm1), len(norm2))

        return common_chars / max_length if max_length > 0 else 0.0
