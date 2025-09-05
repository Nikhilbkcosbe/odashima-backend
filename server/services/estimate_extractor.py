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
        """Extract required information from the specified page of the estimate reference PDF.
        Now extracts: 工種区分, 工事中止日数, 単価地区, 単価使用年月, 歩掛適用年月
        """
        try:
            if len(self.pdf_pages) <= page_index:
                logger.warning(
                    f"PDF has less than {page_index + 1} pages, cannot extract estimate info from page {page_index + 1}")
                return {
                    '工種区分': 'Not Found',
                    '工事中止日数': 'Not Found',
                    '単価地区': 'Not Found',
                    '単価使用年月': 'Not Found',
                    '歩掛適用年月': 'Not Found'
                }

            # Get the specified page
            page = self.pdf_pages[page_index]
            page_text = page.extract_text() or ""
            page_text = self._clean_text(page_text)
        except Exception as e:
            logger.error(f"Error extracting text from PDF page: {str(e)}")
            return {
                '工種区分': 'Not Found',
                '工事中止日数': 'Not Found',
                '単価地区': 'Not Found',
                '単価使用年月': 'Not Found',
                '歩掛適用年月': 'Not Found'
            }

        logger.info(
            f"Extracted text from page {page_index + 1}: {page_text[:500]}...")
        logger.info(f"Full page text length: {len(page_text)}")

        def _normalize_digits(s: str) -> str:
            if not s:
                return s
            trans = str.maketrans({'０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
                                   '５': '5', '６': '6', '７': '7', '８': '8', '９': '9'})
            return s.translate(trans)

        # 工種区分: value immediately after the label
        koushu_patterns = [
            r'工\s*種\s*区\s*分\s*[:：]?\s*([^\s\n]+)',
            r'工種区分[^\n]*?([^\s\n]+)'
        ]
        koushu = 'Not Found'
        for pattern in koushu_patterns:
            m = re.search(pattern, page_text)
            if m and m.group(1):
                koushu = m.group(1).strip()
                break

        # 工事中止日数: capture '<n>日' (or '日間') and normalize to '<n>日'
        chushi = 'Not Found'
        m = re.search(
            r'工\s*事\s*中\s*止\s*日\s*数[^\n]*?([0-9０-９]+)\s*日(?:\s*間)?', page_text)
        if m and m.group(1):
            chushi = _normalize_digits(m.group(1)) + '日'

        # 単価地区
        tanka_chiku = 'Not Found'
        m = re.search(r'単\s*価\s*地\s*区\s*[:：]?\s*([^\s\n]+)', page_text)
        if m and m.group(1):
            tanka_chiku = m.group(1).strip()

        # 単価使用年月: capture 'YYYY年 M月' allowing spaces and full-width digits
        tanka_ym = 'Not Found'
        m = re.search(
            r'単\s*価\s*使\s*用\s*年\s*月\s*[:：]?\s*([0-9０-９]{4})年\s*([0-9０-９]{1,2})月', page_text)
        if m and m.group(1) and m.group(2):
            year = _normalize_digits(m.group(1))
            month = _normalize_digits(m.group(2))
            tanka_ym = f"{year}年 {month}月"
        else:
            m2 = re.search(
                r'単\s*価\s*使\s*用\s*年\s*月\s*[:：]?\s*([^\s\n]+)', page_text)
            if m2 and m2.group(1):
                tanka_ym = m2.group(1).strip()

        # 歩掛適用年月: capture 'YYYY年 M月' allowing spaces and full-width digits
        hokake_ym = 'Not Found'
        m = re.search(
            r'歩\s*掛\s*適\s*用\s*年\s*月\s*[:：]?\s*([0-9０-９]{4})年\s*([0-9０-９]{1,2})月', page_text)
        if m and m.group(1) and m.group(2):
            year = _normalize_digits(m.group(1))
            month = _normalize_digits(m.group(2))
            hokake_ym = f"{year}年 {month}月"
        else:
            m2 = re.search(
                r'歩\s*掛\s*適\s*用\s*年\s*月\s*[:：]?\s*([^\s\n]+)', page_text)
            if m2 and m2.group(1):
                hokake_ym = m2.group(1).strip()

        # 総日数: e.g., ３３５日間 (appears before table). Capture first occurrence on the page
        total_days = 'Not Found'
        m = re.search(r'([0-9０-９]+)\s*日\s*間', page_text)
        if m and m.group(1):
            # Preserve original width of digits
            total_days = m.group(1) + '日間'

        result = {
            '工種区分': koushu,
            '工事中止日数': chushi,
            '単価地区': tanka_chiku,
            '単価使用年月': tanka_ym,
            '歩掛適用年月': hokake_ym,
            '総日数': total_days
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
