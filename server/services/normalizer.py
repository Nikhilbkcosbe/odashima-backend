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
        # Preserve × (multiplication), ✕ (cross), and numerical indicators like ２, ３, etc.
        normalized = re.sub(
            r'[^\w\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF×✕]', '', normalized)

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
        Check if two items are significantly different and should not be considered matches.
        Enhanced to handle PDF (detailed) vs Excel (base name) matching patterns.
        """
        if not text1 or not text2:
            return True

        import re

        # STEP 0: Extract base item names by removing "+ specification" parts
        # This handles cases like "防護柵設置 + 歩行者自転車柵兼用,B種,H=950mm" vs "防護柵設置"
        def extract_base_name(text):
            # Remove everything after " + " to get base item name
            base = re.split(r'\s*\+\s*', text)[0].strip()
            return self._normalize_text(base)

        base1 = extract_base_name(text1)
        base2 = extract_base_name(text2)

        # If base names match, these are likely the same item (PDF detailed vs Excel base)
        if base1 == base2 and base1:
            return False

        # If base names are very similar (fuzzy match), they're likely the same item
        if base1 and base2:
            # Simple character-based similarity for base names
            common_chars = sum(1 for c in base1 if c in base2)
            max_length = max(len(base1), len(base2))
            similarity = common_chars / max_length if max_length > 0 else 0.0

            # If base names are >85% similar, consider them the same item
            if similarity > 0.85:
                return False

        # STEP 1: Normalize both full texts for detailed comparison
        norm1 = self._normalize_text(text1)
        norm2 = self._normalize_text(text2)

        # STEP 2: Check for specific patterns that indicate different items

        # 2a. Different numerical suffixes (e.g., ベンチフリュームボックス vs ベンチフリュームボックス２)
        # Extract base names (without numbers)
        base_no_nums1 = re.sub(r'[０-９0-9]+', '', norm1)
        base_no_nums2 = re.sub(r'[０-９0-9]+', '', norm2)

        # If base names are the same but original texts have different numbers, they're different items
        if base_no_nums1 == base_no_nums2 and base_no_nums1:
            # Extract all numbers from both texts
            numbers1 = re.findall(r'[０-９0-9]+', norm1)
            numbers2 = re.findall(r'[０-９0-9]+', norm2)

            # If they have different numbers, consider them different
            if numbers1 != numbers2:
                return True

        # 2b. Different multiplication factors (e.g., 800×590×2000 vs 800×590×2000✕2)
        # Check for ✕ followed by numbers
        mult_pattern1 = re.search(r'✕([０-９0-9]+)', norm1)
        mult_pattern2 = re.search(r'✕([０-９0-9]+)', norm2)

        # If one has multiplication factor and other doesn't, they're different
        if bool(mult_pattern1) != bool(mult_pattern2):
            return True

        # If both have multiplication factors but different values, they're different
        if mult_pattern1 and mult_pattern2:
            if mult_pattern1.group(1) != mult_pattern2.group(1):
                return True

        # STEP 3: Significant length difference (but more lenient for base name matches)
        # Only apply strict length checking if base names don't match at all
        if len(norm1) > 0 and len(norm2) > 0:
            length_ratio = abs(len(norm1) - len(norm2)) / \
                max(len(norm1), len(norm2))

            # If base names are completely different AND length difference >50%, consider different
            # (More lenient than the original 30% to allow PDF detailed vs Excel base matching)
            if length_ratio > 0.5:
                return True

        return False

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
