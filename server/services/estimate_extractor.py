import os
import re
import pdfplumber
import logging

logger = logging.getLogger(__name__)


class EstimateReferenceExtractor:
    def __init__(self, pdf_path: str):
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF not found: {pdf_path}")
        self.pdf_path = pdf_path
        self.pdf_pages = []
        self._load_pdf()

    def _load_pdf(self):
        """Load the PDF and extract pages"""
        try:
            # Keep the PDF object open to prevent file handle issues
            self.pdf = pdfplumber.open(self.pdf_path)
            self.pdf_pages = self.pdf.pages
        except Exception as e:
            logger.error(f"Error loading PDF: {str(e)}")
            raise

    def __del__(self):
        """Cleanup method to close PDF file"""
        try:
            if hasattr(self, 'pdf') and self.pdf:
                self.pdf.close()
        except Exception:
            pass

    def _search(self, pattern, text, group=1):
        """Search for pattern in text and return the specified group"""
        if not text:
            return "Not Found"
        match = re.search(pattern, text, re.DOTALL | re.MULTILINE)
        return (match.group(group) or "").strip() if match else "Not Found"

    def _clean_text(self, text):
        """Clean text by removing extra spaces and normalizing characters"""
        if not text:
            return ""
        # Remove full-width and half-width spaces, normalize spaces
        text = re.sub(r'[\u3000\u0020]+', ' ', text)
        # Remove newlines and tabs
        text = re.sub(r'[\n\r\t]+', ' ', text)
        # Remove multiple spaces
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    def extract_estimate_info(self, page_index=1):
        """Extract required information from the specified page of the estimate reference PDF"""
        try:
            if len(self.pdf_pages) <= page_index:
                logger.warning(
                    f"PDF has less than {page_index + 1} pages, cannot extract estimate info from page {page_index + 1}")
                return {
                    '省庁': 'Not Found',
                    '年度': 'Not Found',
                    '経費工種': 'Not Found',
                    '施工地域工事場所': 'Not Found'
                }

            # Get the specified page
            page = self.pdf_pages[page_index]
            page_text = page.extract_text() or ""
            page_text = self._clean_text(page_text)
        except Exception as e:
            logger.error(f"Error extracting text from PDF page: {str(e)}")
            return {
                '省庁': 'Not Found',
                '年度': 'Not Found',
                '経費工種': 'Not Found',
                '施工地域工事場所': 'Not Found'
            }

        logger.info(f"Extracted text from page {page_index + 1}: {page_text[:500]}...")
        logger.info(f"Full page text length: {len(page_text)}")

        # Extract 省庁 from footer position (usually at the bottom of the page)
        # Look for patterns that include "県" or "都" or "府" or "市"
        shocho_patterns = [
            r'([^\s]+[県都府市])',  # Any text ending with 県, 都, 府, 市
            r'([^\s]+県[^\s]*)',    # Text containing 県
            r'([^\s]+都[^\s]*)',    # Text containing 都
            r'([^\s]+府[^\s]*)',    # Text containing 府
            r'([^\s]+市[^\s]*)',    # Text containing 市
        ]

        shocho = "Not Found"
        for pattern in shocho_patterns:
            matches = re.findall(pattern, page_text)
            if matches:
                # Take the last match (usually at the bottom/footer)
                shocho = matches[-1]
                break

        # If still not found, try to extract from the end of the text (footer area)
        if shocho == "Not Found":
            # Get the last 1000 characters (footer area)
            footer_text = page_text[-1000:] if len(
                page_text) > 1000 else page_text
            for pattern in shocho_patterns:
                matches = re.findall(pattern, footer_text)
                if matches:
                    shocho = matches[-1]
                    break

        # If still not found, try to extract from the very end (last 500 characters)
        if shocho == "Not Found":
            # Get the last 500 characters (very bottom of page)
            bottom_text = page_text[-500:] if len(
                page_text) > 500 else page_text
            for pattern in shocho_patterns:
                matches = re.findall(pattern, bottom_text)
                if matches:
                    shocho = matches[-1]
                    break

        # If still not found, try to look for common prefecture names
        if shocho == "Not Found":
            common_prefectures = [
                r'(東京都)', r'(神奈川県)', r'(千葉県)', r'(埼玉県)', r'(茨城県)',
                r'(栃木県)', r'(群馬県)', r'(山梨県)', r'(静岡県)', r'(愛知県)',
                r'(三重県)', r'(滋賀県)', r'(京都府)', r'(大阪府)', r'(兵庫県)',
                r'(奈良県)', r'(和歌山県)', r'(岐阜県)', r'(富山県)', r'(石川県)',
                r'(福井県)', r'(新潟県)', r'(長野県)', r'(山形県)', r'(福島県)',
                r'(宮城県)', r'(秋田県)', r'(青森県)', r'(岩手県)', r'(北海道)',
            ]
            for pattern in common_prefectures:
                match = self._search(pattern, page_text)
                if match != "Not Found":
                    shocho = match
                    break

                # Extract 年度 (Japanese era + number + 年度)
        nendo_patterns = [
            r'([令和平成昭和大正明治]\s*\d+\s*年度)',  # Full pattern with era
            r'(\d+\s*年度)',                           # Just number + 年度
            r'([令和平成昭和大正明治]\s*\d+)',         # Era + number
            # Just number + 年度 (no spaces)
            r'(\d+年度)',
        ]

        nendo = "Not Found"
        for pattern in nendo_patterns:
            match = self._search(pattern, page_text)
            if match != "Not Found":
                nendo = match
                break

        # If still not found, try to extract from the beginning of the text (header area)
        if nendo == "Not Found":
            # Get the first 1000 characters (header area)
            header_text = page_text[:1000] if len(
                page_text) > 1000 else page_text
            for pattern in nendo_patterns:
                match = self._search(pattern, header_text)
                if match != "Not Found":
                    nendo = match
                    break

        # If still not found, try to extract from the very beginning (first 500 characters)
        if nendo == "Not Found":
            # Get the first 500 characters (very top of page)
            top_text = page_text[:500] if len(page_text) > 500 else page_text
            for pattern in nendo_patterns:
                match = self._search(pattern, top_text)
                if match != "Not Found":
                    nendo = match
                    break

        # If still not found, try to look for common year patterns
        if nendo == "Not Found":
            # Look for any year pattern in the text
            year_patterns = [
                r'(令和\d+年度)', r'(平成\d+年度)', r'(昭和\d+年度)',
                r'(令和\d+)', r'(平成\d+)', r'(昭和\d+)',
                r'(\d+年度)', r'(\d+年)',
            ]
            for pattern in year_patterns:
                match = self._search(pattern, page_text)
                if match != "Not Found":
                    nendo = match
                    break

                # Extract 経費工種 (comes after 工種区分)
        # Look for the value in the 摘要 column after 工種区分
        keihi_patterns = [
            r'工種区分[^\n]*?([^\s\n]+)',  # After 工種区分, get the next word
            r'工種区分\s*([^\s\n]+)',      # Directly after 工種区分, get next word
            # After 工種区分, get text until space or end
            r'工種区分[^\n]*?([^\n]+?)(?=\s|$)',
        ]

        keihi = "Not Found"
        for pattern in keihi_patterns:
            match = self._search(pattern, page_text)
            if match != "Not Found":
                # Clean up the extracted value - take only the first word/phrase
                keihi = match.split()[0] if match else match
                break

        # If still not found, try to find 土木 or 橋梁 in the text
        if keihi == "Not Found":
            fallback_patterns = [
                r'([^\s\n]*土木[^\s\n]*)',  # Look for 土木
                r'([^\s\n]*橋梁[^\s\n]*)',  # Look for 橋梁
            ]
            for pattern in fallback_patterns:
                match = self._search(pattern, page_text)
                if match != "Not Found":
                    keihi = match
                    break

                # Extract 施工地域工事場所 (comes after 工事名)
        # Based on the actual text, it seems to be "一般国道107号水沢橋橋梁補修その２工事"
        kouji_patterns = [
            # After 工事名, get text until space or end
            r'工\s*事\s*名[^\n]*?([^\n]+?)(?=\s|$)',
            # Directly after 工事名, get next word
            r'工\s*事\s*名\s*([^\s\n]+)',
            # Pattern for highway construction (word boundary)
            r'(一般国道\d+号[^\s\n]+工事)',
            # Pattern for bridge construction (word boundary)
            r'([^\s\n]*橋[^\s\n]*工事)',
            # Highway number pattern (word boundary)
            r'(一般国道\d+号[^\s\n]+)',
            # Specific bridge name (word boundary)
            r'([^\s\n]*水沢橋[^\s\n]*)',
        ]

        kouji = "Not Found"
        for pattern in kouji_patterns:
            match = self._search(pattern, page_text)
            if match != "Not Found":
                # Clean up the extracted value - take only the construction name
                kouji = match.split()[0] if match else match
                break

        # If still not found, try to extract from the beginning of the text
        if kouji == "Not Found":
            # Look for the first occurrence of construction-related text
            construction_patterns = [
                r'(一般国道\d+号[^\s\n]+)',
                r'([^\s\n]*橋[^\s\n]*工事)',
                r'([^\s\n]*補修[^\s\n]*)',
            ]
            for pattern in construction_patterns:
                match = self._search(pattern, page_text)
                if match != "Not Found":
                    kouji = match.split()[0] if match else match
                    break

        result = {
            '省庁': shocho,
            '年度': nendo,
            '経費工種': keihi,
            '施工地域工事場所': kouji
        }

        logger.info(f"Extracted estimate info: {result}")
        return result

    def close(self):
        """Explicitly close the PDF file"""
        try:
            if hasattr(self, 'pdf') and self.pdf:
                self.pdf.close()
                self.pdf = None
        except Exception as e:
            logger.error(f"Error closing PDF: {str(e)}")
