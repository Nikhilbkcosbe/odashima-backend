import os
import re
import pdfplumber
import json
import pandas as pd


class SpecFinalExtractor:
    def __init__(self, pdf_path: str):
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF not found: {pdf_path}")
        self.pdf_path = pdf_path
        self.article_map = {}
        self.pdf_pages = []
        self._index_pdf()

    def _index_pdf(self):
        with pdfplumber.open(self.pdf_path) as pdf:
            self.pdf_pages = pdf.pages
            pages_words = [page.extract_words(
                x_tolerance=1, y_tolerance=1) for page in self.pdf_pages]

        article_pattern = re.compile(r"(第(\d+)条)")
        raw_found = []
        for i, words in enumerate(pages_words):
            if not words:
                continue
            for word in words:
                if word['x0'] < 100 and word['top'] < 150:
                    match = article_pattern.match(word['text'])
                    if match:
                        article_name = match.group(1)
                        article_num = int(match.group(2))
                        if not any(a['name'] == article_name for a in raw_found):
                            raw_found.append(
                                {'name': article_name, 'num': article_num, 'page': i})
                        break

        sorted_articles = sorted(raw_found, key=lambda x: x['num'])
        self.article_map = {}
        for i, article in enumerate(sorted_articles):
            start_page = article['page']
            end_page = len(self.pdf_pages)
            if i + 1 < len(sorted_articles):
                end_page = sorted_articles[i+1]['page']
            self.article_map[article['name']] = {
                'start': start_page, 'end': end_page}

    def _get_content_for_article(self, article_name):
        if article_name not in self.article_map:
            return "", []
        map_info = self.article_map[article_name]
        start_page, end_page = map_info['start'], map_info['end']
        text_content = ""
        all_tables = []
        for i in range(start_page, end_page):
            page = self.pdf_pages[i]
            text_content += page.extract_text(x_tolerance=1) or ""
            extracted_tables = page.extract_tables()
            if extracted_tables:
                all_tables.extend(extracted_tables)
        return text_content, all_tables

    def _search(self, pattern, text, group=1):
        if not text:
            return "Not Found"
        match = re.search(pattern, text, re.DOTALL | re.MULTILINE)
        return (match.group(group) or "").strip() if match else "Not Found"

    # The following extraction methods are adapted from final_extractor.py
    def extract_dai2jou(self):
        text, _ = self._get_content_for_article("第２条")
        yoyuu_kikan = self._search(r"うち余裕期間\s*(\d*)\s*日間", text) or ""
        jitsu_kouki = self._search(r"うち実工期\s*(\d*)\s*日間", text) or ""
        results = {
            '全体工期': self._search(r"全体工期\s+(\d+)", text),
            'うち余裕期間': yoyuu_kikan,
            'うち実工期': jitsu_kouki,
            '余裕期間の設定_対象の有無': self._search(r"3\s+余裕期間の設定.*?対象の有無\s+([^\n]+)", text),
            '週休２日工事_対象の有無': self._search(r"4\s+週休２日工事.*?対象の有無\s+(.+)", text),
            '週休２日工事_分類': self._search(r"4\s+週休２日工事.*?\s*（(.*?)）", text),
            '熱中症予防対策に係る工期の延長_対象の有無': self._search(r"11\s+熱中症予防対策.*?対象の有無\s+(.+)", text),
        }
        for key, value in results.items():
            if "対象の有無" in key and value not in ["Not Found", "無", "有"]:
                results[key] = "有" if "有" in value else "無" if "無" in value else "Not Found"
        return results

    def extract_dai3jou(self):
        text, _ = self._get_content_for_article("第３条")
        results = {
            '工事現場の現場環境改善及び地域連携_対象の有無': self._search(r"4\s+工事現場の現場環境改善.*?対象の有無\s+([^\n]+)", text),
            'ＩＣＴ活用工事_対象の有無': self._search(r"15\s+ＩＣＴ活用工事.*?対象の有無\s+([^\n]+)", text),
            'ＩＣＴ活用工事_区分': self._search(r"15\s+ＩＣＴ活用工事.*?\nＩＣＴ活用工事（(.*?)）", text),
            '１日未満で完了する小規模作業の積算_対象の有無': self._search(r"16\s+１日未満で完了する小規模作業の積算.*?対象の有無\s+([^\n]+)", text),
            '熱中症対策に資する現場管理費補_対象の有無': self._search(r"17\s+熱中症対策に資する現場管理費補.*?対象の有無\s+([^\n]+)", text)
        }
        for key, value in results.items():
            if "対象の有無" in key and value not in ["Not Found"]:
                results[key] = "有" if "有" in value else "無"
        return results

    def extract_dai4jou(self):
        text, all_tables = self._get_content_for_article("第４条")
        results = {
            'レディーミクストコンクリート': self._extract_concrete_table(all_tables),
            'アスファルト混合物': self._extract_asphalt_table(all_tables),
            '上記以外の使用アスファルト合材の有無': self._extract_other_asphalt_table(all_tables),
            '石材類': self._extract_stone_table(all_tables),
            '上記以外の使用材料の有無': self._extract_other_materials_table(all_tables),
            '鉄筋': self._extract_rebar_table(all_tables),
            'その他': self._extract_other_table(all_tables)
        }
        return results

    def _extract_concrete_table(self, tables):
        concrete_list = []
        for table in tables:
            if not table:
                continue
            table_text = "".join("".join(filter(
                None, cell)) for row in table for cell in row if cell).replace('\n', '').replace(' ', '')
            if "適用工種" in table_text and "セメント種類" in table_text:
                sub_header_row_idx = -1
                for idx, row in enumerate(table):
                    if row and any(c and ("BB" in c or "N" in c) for c in row):
                        sub_header_row_idx = idx
                        break
                if sub_header_row_idx == -1:
                    continue
                sub_headers = table[sub_header_row_idx]
                for row_idx in range(sub_header_row_idx + 1, len(table)):
                    row = table[row_idx]
                    if row and len(row) >= 8 and '■' in str(row[0]):
                        cleaned = [str(c).replace('\n', '')
                                   if c else '' for c in row]
                        c_type = ""
                        if len(sub_headers) > 3 and '■' in cleaned[3]:
                            c_type = sub_headers[3]
                        elif len(sub_headers) > 4 and '■' in cleaned[4]:
                            c_type = sub_headers[4]
                        concrete_list.append({
                            "セメント種類": c_type.strip(),
                            "規格": cleaned[5],
                            "最大水セメント比": cleaned[6],
                            "最小セメント使用量": cleaned[7]
                        })
        return concrete_list if concrete_list else "Not Found"

    def _extract_asphalt_table(self, tables):
        asphalt_list = []
        for table in tables:
            if not table or not table[0]:
                continue
            table_text = "".join(filter(None, table[0])).replace(
                '\n', '').replace(' ', '')
            if "アスファルト合材名" in table_text and "使用箇所" in table_text and len(table[0]) > 2:
                for row in table[1:]:
                    if row and len(row) >= 4 and '■' in str(row[0]):
                        cleaned = [str(c).replace('\n', '')
                                   if c else '' for c in row]
                        asphalt_list.append(
                            {"材料名": cleaned[2], "使用箇所": cleaned[3]})
        return asphalt_list if asphalt_list else "Not Found"

    def _extract_other_asphalt_table(self, tables):
        for table in tables:
            if not table:
                continue
            table_text = "".join("".join(filter(None, cell))
                                 for row in table for cell in row if cell)
            if "上記以外の使用アスファルト合材の有無" in table_text:
                other_asphalt_list = []
                for row in table:
                    if row and len(row) >= 3 and '■' in str(row[0]):
                        cleaned = [str(c).replace('\n', '')
                                   if c else '' for c in row]
                        other_asphalt_list.append(
                            {"材料名": cleaned[1], "使用箇所": cleaned[2]})
                if other_asphalt_list:
                    return other_asphalt_list
        return "Not Found"

    def _extract_stone_table(self, tables):
        stone_list = []
        for table in tables:
            if not table or not table[0]:
                continue
            header_row_text = "".join(filter(None, table[0])).replace(
                '\n', '').replace(' ', '')
            if "石材類" in header_row_text or ("材料名" in header_row_text and "適用箇所" in header_row_text):
                for row in table[1:]:
                    if row and len(row) >= 4 and '■' in str(row[0]):
                        cleaned = [str(c).replace('\n', ' ')
                                   if c else '' for c in row]
                        stone_list.append(
                            {"材料名": cleaned[1], "規格": cleaned[2], "適用箇所": cleaned[3]})
        return stone_list if stone_list else "Not Found"

    def _extract_other_materials_table(self, tables):
        other_materials_list = []
        for table in tables:
            if not table or not table[0]:
                continue
            table_text = "".join("".join(filter(None, cell))
                                 for row in table for cell in row if cell)
            if "上記以外の使用材料の有無" in table_text:
                for row in table:
                    if row and len(row) >= 4 and '■' in str(row[0]):
                        cleaned = [str(c).replace('\n', ' ')
                                   if c else '' for c in row]
                        other_materials_list.append(
                            {"材料名": cleaned[1], "規格": cleaned[2], "適用箇所": cleaned[3]})
        return other_materials_list if other_materials_list else "Not Found"

    def _extract_rebar_table(self, tables):
        rebar_list = []
        for table in tables:
            if not table or not table[0]:
                continue
            header_row_text = "".join(filter(None, table[0])).replace(
                '\n', '').replace(' ', '')
            if "鉄筋" in header_row_text:
                for row in table[1:]:
                    if row and len(row) >= 4 and '■' in str(row[0]):
                        cleaned = [str(c).replace('\n', ' ')
                                   if c else '' for c in row]
                        rebar_list.append(
                            {"材料名": cleaned[1], "規格": cleaned[2], "適用箇所": cleaned[3]})
        return rebar_list if rebar_list else "Not Found"

    def _extract_other_table(self, all_tables):
        found_keyword_table_idx = -1
        for i, table in enumerate(all_tables):
            table_text = "".join(str(cell)
                                 for row in table for cell in row if cell)
            if "その他の使用材料の有無" in table_text:
                found_keyword_table_idx = i
                break
        if found_keyword_table_idx != -1:
            for j in range(found_keyword_table_idx, len(all_tables)):
                table = all_tables[j]
                if table and table[0]:
                    header = table[0]
                    cleaned_header = [str(h).replace(
                        '\n', '').strip() for h in header]
                    if '材料名' in cleaned_header and '規格・寸法・材質' in cleaned_header:
                        data = []
                        for row in table[1:]:
                            if any(cell and str(cell).strip() for cell in row):
                                cleaned_row = [str(c).replace(
                                    '\n', ' ').strip() if c else '' for c in row]
                                data.append({
                                    "材料名": cleaned_row[0] if len(cleaned_row) > 0 else '',
                                    "規格・寸法・材質": cleaned_row[1] if len(cleaned_row) > 1 else '',
                                    "適用工種": cleaned_row[2] if len(cleaned_row) > 2 else '',
                                    "備考": cleaned_row[3] if len(cleaned_row) > 3 else ''
                                })
                        return data if data else "Not Found"
        return "Not Found"

    def extract_dai7jou(self):
        text, _ = self._get_content_for_article("第７条")
        lines = text.split('\n')
        results = {
            '公害防止のための制限': "Not Found",
            '水替・流入防止施設': "Not Found",
            '濁水・湧水等の処理条件': "Not Found",
            '事業損失防止': "Not Found"
        }
        for i, line in enumerate(lines):
            if '排出ガス防止のための施工方法等の制限の有無' in line and '有' in line:
                results['公害防止のための制限'] = '排出ガス防止のための施工方法等の制限の有無'
            if '水替・流入防止施設設置の公害防止対策の有無' in line and '有' in line:
                facility_info = {}
                if i + 1 < len(lines) and '施 設 内 容' in lines[i+1]:
                    value = lines[i+1].replace('施 設 内 容', '').strip()
                    if value:
                        facility_info['施設内容'] = value
                if i + 2 < len(lines) and '設 置 期 間' in lines[i+2]:
                    value = lines[i+2].replace('設 置 期 間', '').strip()
                    if value:
                        facility_info['設置期間'] = value
                if facility_info:
                    results['水替・流入防止施設'] = facility_info
            if '濁水・湧水等の処理条件の有無' in line and '有' in line:
                treatment_info = {}
                if i + 1 < len(lines) and '処 理 施 設' in lines[i+1]:
                    value = lines[i+1].replace('処 理 施 設', '').strip()
                    if value:
                        treatment_info['処理施設'] = value
                if i + 2 < len(lines) and '処 理 条 件 等' in lines[i+2]:
                    value = lines[i+2].replace('処 理 条 件 等', '').strip()
                    if value:
                        treatment_info['処理条件等'] = value
                if treatment_info:
                    results['濁水・湧水等の処理条件'] = treatment_info
            if '事業損失防止のための事前・事後調査の有無' in line and '有' in line:
                loss_prevention_info = {}
                if i + 1 < len(lines) and '調 査 項 目' in lines[i+1]:
                    value = lines[i+1].replace('調 査 項 目', '').strip()
                    if value:
                        loss_prevention_info['調査項目'] = value
                if i + 2 < len(lines) and '調 査 時 期' in lines[i+2]:
                    value = lines[i+2].replace('調 査 時 期', '').strip()
                    if value:
                        loss_prevention_info['調査時期'] = value
                if i + 3 < len(lines) and '調 査 方 法' in lines[i+3]:
                    value = lines[i+3].replace('調 査 方 法', '').strip()
                    if value:
                        loss_prevention_info['調査方法'] = value
                if i + 4 < len(lines) and '調 査 範 囲' in lines[i+4]:
                    value = lines[i+4].replace('調 査 範 囲', '').strip()
                    if value:
                        loss_prevention_info['調査範囲'] = value
                if loss_prevention_info:
                    results['事業損失防止'] = loss_prevention_info
        return results

    def extract_dai8jou(self):
        results = {"交通誘導警備員": "Not Found"}
        _, all_tables = self._get_content_for_article("第８条")
        target_headers = ['配置場所', '配置員数', '編制', '総配置員数', '昼夜別', '交代要員の有無']
        for table in all_tables:
            if not table or not table[0]:
                continue
            header_row_text = "".join(filter(None, table[0]))
            if all(h in header_row_text for h in target_headers):
                if len(table) > 1:
                    data_row = table[1]
                    if len(data_row) == len(target_headers):
                        haichi_basho_raw = data_row[0] or ""
                        rosenmei = ""
                        if '\n' in haichi_basho_raw:
                            rosenmei = haichi_basho_raw.split('\n')[1]
                        extracted_data = {
                            "路線名": rosenmei,
                            "配置員数": (data_row[1] or "").strip(),
                            "編制": (data_row[2] or "").strip(),
                            "総配置員数": (data_row[3] or "").strip(),
                            "昼夜別": (data_row[4] or "").strip(),
                            "交代要員の有無": (data_row[5] or "").strip(),
                        }
                        results["交通誘導警備員"] = extracted_data
                        break
        return results

    def extract_dai10jou(self):
        text, tables = self._get_content_for_article("第10条")
        results = {}
        kasetsu_headers = ['工種', '種別', '細別', '単位', '数量', '備考']
        kasetsu_tables = []
        for table in tables:
            if table and table[0]:
                header_text = "".join(str(c)
                                      for c in table[0]).replace('\n', '')
                if all(h in header_text for h in kasetsu_headers):
                    kasetsu_tables.append(table)
        if len(kasetsu_tables) > 0:
            data = [dict(zip(kasetsu_headers, [str(c or '').strip() for c in row]))
                    for row in kasetsu_tables[0][1:] if any(c and str(c).strip() for c in row)]
            results['任意仮設'] = data if data else "Not Found"
        else:
            results['任意仮設'] = "Not Found"
        if len(kasetsu_tables) > 1:
            data = [dict(zip(kasetsu_headers, [str(c or '').strip() for c in row]))
                    for row in kasetsu_tables[1][1:] if any(c and str(c).strip() for c in row)]
            results['指定仮設'] = data if data else "Not Found"
        else:
            results['指定仮設'] = "Not Found"
        paired_items_config = [
            ("仮設備の引渡し・引継ぎ", "仮設備の引渡し・引継ぎの有無", "仮設備の引渡し・引継ぎ詳細"),
            ("仮設備の構造・施工方法の指定", "仮設備の構造・施工方法の指定の有無", "仮設備の構造・施工方法の指定詳細"),
            ("仮設備の設計条件の指定", "仮設備の設計条件の指定の有無", "仮設備の設計条件の指定詳細")
        ]
        detail_tables = [tbl for tbl in tables if tbl and len(
            tbl[0]) == 2 and "対象の有無" not in str(tbl)]
        for i, (keyword, key_yes_no, key_details) in enumerate(paired_items_config):
            keyword_pos = text.find(keyword)
            if keyword_pos != -1:
                search_slice = text[keyword_pos:keyword_pos + 75]
                if '無' in search_slice:
                    results[key_yes_no] = '無'
                elif '有' in search_slice:
                    results[key_yes_no] = '有'
                else:
                    results[key_yes_no] = 'Not Found'
            else:
                results[key_yes_no] = 'Not Found'
            if i < len(detail_tables):
                table_data = []
                for row in detail_tables[i]:
                    if len(row) == 2 and all(c and str(c).strip() for c in row):
                        table_data.append([str(c).strip() for c in row])
                results[key_details] = table_data if table_data else "Not Found"
            else:
                results[key_details] = "Not Found"
        return results

    def extract_dai11jou(self):
        text, tables = self._get_content_for_article("第11条")
        results = {}
        results['土捨て場'] = "Not Found"
        fukusanbutsu_headers = ['副産物名', '搬入再資源化施設名', '搬入場所', '備考']
        fukusanbutsu_data = []
        found_fukusanbutsu = False
        for table in tables:
            if not table or not table[0]:
                continue
            header_text = "".join(str(c or '').replace('\n', '')
                                  for c in table[0])
            if '副産物名' in header_text and '搬入再資源化施設名' in header_text:
                for row in table[1:]:
                    if row and len(row) >= 3 and all(str(c).strip() for c in row[:3]):
                        padded_row = (
                            list(row) + [None] * len(fukusanbutsu_headers))[:len(fukusanbutsu_headers)]
                        fukusanbutsu_data.append(
                            dict(zip(fukusanbutsu_headers, [str(c or '').strip() for c in padded_row])))
                if fukusanbutsu_data:
                    found_fukusanbutsu = True
                    break
        results['建設副産物'] = fukusanbutsu_data if fukusanbutsu_data else "Not Found"
        haikibutsu_headers = ['廃棄物名', '受入施設名', '受入場所', '備考']
        haikibutsu_data = []
        for table in tables:
            if not table or not table[0]:
                continue
            header_text = "".join(str(c or '').replace('\n', '')
                                  for c in table[0])
            if '廃棄物名' in header_text and '受入施設名' in header_text:
                for row in table[1:]:
                    if row and len(row) >= 3 and all(str(c).strip() for c in row[:3]):
                        padded_row = (
                            list(row) + [None] * len(haikibutsu_headers))[:len(haikibutsu_headers)]
                        haikibutsu_data.append(
                            dict(zip(haikibutsu_headers, [str(c or '').strip() for c in padded_row])))
                if haikibutsu_data:
                    break
                results['建設廃棄物'] = haikibutsu_data if haikibutsu_data else "Not Found"
        return results

    def extract_dai13jou(self):
        text, tables = self._get_content_for_article("第13条")
        results = {}
        yakueki_umu = self._search(r"薬液注入を行う場合.*?対象の有無\s*([^\n]+)", text)
        yakueki_details = []
        if "有" in yakueki_umu:
            pass
        results['薬液注入を行う場合'] = yakueki_details if yakueki_details else "Not Found"
        shuuhen_umu = self._search(r"周辺環境影響調査.*?対象の有無\s*([^\n]+)", text)
        shuuhen_details = []
        if "有" in shuuhen_umu:
            chousa_headers = ['調査項目', '採取地点', '採取回数', '備考']
            for table in tables:
                if not table or not table[0]:
                    continue
                header_text = "".join(str(c or '').replace('\n', '')
                                      for c in table[0])
                if all(h in header_text for h in chousa_headers):
                    for row in table[1:]:
                        if row and len(row) >= 2 and all(str(c).strip() for c in row[:2]):
                            padded_row = (
                                list(row) + [None] * len(chousa_headers))[:len(chousa_headers)]
                            shuuhen_details.append(
                                dict(zip(chousa_headers, [str(c or '').strip() for c in padded_row])))
                    if shuuhen_details:
                        break
        results['周辺環境影響調査'] = shuuhen_details if shuuhen_details else "Not Found"
        return results

    def extract_dai14jou(self):
        text, tables = self._get_content_for_article("第14条")
        results = {}
        hasseihin_headers = ['種類', '数量', '保管・仮置場所']
        hasseihin_data = []
        for table in tables:
            if not table or not table[0]:
                continue
            header_text = "".join(str(c or '').replace(
                '\n', '').replace(' ', '') for c in table[0])
            if all(h in header_text for h in hasseihin_headers):
                for row in table[1:]:
                    if row and len(row) >= 3 and all(str(c).strip() for c in row[:3]):
                        padded_row = (
                            list(row) + [None] * len(hasseihin_headers))[:len(hasseihin_headers)]
                        hasseihin_data.append(
                            dict(zip(hasseihin_headers, [str(c or '').strip() for c in padded_row])))
                if hasseihin_data:
                    break
        results['現場発生品'] = hasseihin_data if hasseihin_data else "Not Found"

        umu_key = '労働者確保に要する間接費の実績変更_対象の有無'
        results[umu_key] = self._search(
            r"労働者確保に要する間接費の実績変更.*?対象の有無\s*([^\n]+)", text)
        if results[umu_key] not in ["有", "無"]:
            results[umu_key] = "有" if "有" in (results[umu_key] or "") else "無"

        details_key = '労働者確保に要する間接費の実績変更_詳細'
        roudousha_details = []
        found_detail = False
        for table in tables:
            if found_detail:
                break
            is_target_table = False
            for row in table:
                if len(row) == 2 and str(row[0]).strip() == '○':
                    is_target_table = True
                    break
            if is_target_table:
                for row in table:
                    if len(row) == 2 and str(row[0]).strip() == '○':
                        detail_text = str(row[1] or '').replace(
                            '\n', ' ').strip()
                        if detail_text:
                            roudousha_details.append(detail_text)
                        found_detail = True
                        break
        results[details_key] = roudousha_details if roudousha_details else "Not Found"

        sekou_key_umu = '施工箇所が点在する工事の積算方法_対象の有無'
        sekou_umu = self._search(r"施工箇所が点在する工事の積算方法.*?対象の有無\s*([^\n]+)", text)
        results[sekou_key_umu] = "有" if "有" in sekou_umu else "無"
        if results[sekou_key_umu] == "有":
            sekou_key_details = '施工箇所が点在する工事の積算方法'
            details_text = self._search(
                r"施工箇所が点在する工事の積算方法\s*([^\n]+)\n\s*対象の有無", text)
            results[sekou_key_details] = details_text if details_text else "Not Found"

        tokki_headers = ['特記事項', '特記事項の内容']
        tokki_data = []
        for table in tables:
            if not table or not table[0]:
                continue
            header_text = "".join(str(c or '').replace('\n', '')
                                  for c in table[0])
            if all(h in header_text for h in tokki_headers):
                for row in table[1:]:
                    if row and len(row) >= 2 and all(str(c).strip() for c in row[:2]):
                        padded_row = (
                            list(row) + [None] * len(tokki_headers))[:len(tokki_headers)]
                        tokki_data.append(
                            dict(zip(tokki_headers, [str(c or '').strip() for c in padded_row])))
                if tokki_data:
                    break
        results['その他の特記事項'] = tokki_data if tokki_data else "Not Found"
        return results

    def extract_all(self):
        results = [
            ("第２条", self.extract_dai2jou()),
            ("第３条", self.extract_dai3jou()),
            ("第４条", self.extract_dai4jou()),
            ("第７条", self.extract_dai7jou()),
            ("第８条", self.extract_dai8jou()),
            ("第10条", self.extract_dai10jou()),
            ("第１１条", self.extract_dai11jou()),
            ("第13条", self.extract_dai13jou()),
            ("第14条", self.extract_dai14jou())
        ]
        return results
