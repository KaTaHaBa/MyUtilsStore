import os
import glob
import zipfile
import json
import re
import time
from datetime import datetime
from bs4 import BeautifulSoup
import pandas as pd

# =============================================================================
# 定数設定
# =============================================================================
DEFAULT_BASE_CACHE_NAME = "taxonomy_base_cache.json"

# =============================================================================
# 統合実行関数
# =============================================================================
def execute_process(
    xbrl_zip_dir,
    taxonomy_dir,
    map_cache_dir,
    target_companies,
    target_years,
    rebuild_map=False,
    output_base_dir="out"
):
    """
    EDINET解析の一連のフローを一括実行するラッパー関数
    """
    print("="*60)
    print("EDINET XBRL PARSER - AUTOMATED PROCESS START")
    print("="*60)

    # 1. 初期化と汎用タクソノミのロード
    tm = TaxonomyManager(taxonomy_dir, map_cache_dir)
    tm.load_base_taxonomy()

    # 2. 対象Zipファイルの収集
    if not os.path.exists(xbrl_zip_dir):
        print(f"【エラー】XBRLディレクトリが見つかりません: {xbrl_zip_dir}")
        return pd.DataFrame()
        
    xbrl_zips = glob.glob(os.path.join(xbrl_zip_dir, "*.zip"))
    if not xbrl_zips:
        print("【警告】対象ディレクトリに .zip ファイルがありません。")
        return pd.DataFrame()

    # 3. 企業別拡張タクソノミの解析・キャッシュ
    tm.build_company_specific_caches(
        xbrl_zip_paths=xbrl_zips,
        target_companies=target_companies,
        rebuild=rebuild_map
    )

    # 4. データ抽出実行
    extractor = XbrlExtractor(tm)
    result_data = extractor.extract(
        zip_files=xbrl_zips,
        target_patterns=target_companies,
        target_years=target_years
    )

    # 5. CSV保存と結果返却
    if not result_data:
        print("\n【結果】条件に合致するデータは見つかりませんでした。")
        return pd.DataFrame()

    df = pd.DataFrame(result_data)
    
    # 出力フォルダ作成
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    save_dir = os.path.join(output_base_dir, f"result-{timestamp}")
    os.makedirs(save_dir, exist_ok=True)
    print(f"\n[Output] CSV出力先: {save_dir}")

    count_files = 0
    for (company, code), group_df in df.groupby(['企業名', '証券コード']):
        # 整形（ソート）
        cat_order = {'BS': 0, 'PL': 1, 'CF': 2, 'SS': 3, 'Notes': 4, 'Others': 5}
        group_df['sort_key'] = group_df['カテゴリ'].map(cat_order).fillna(99)
        
        out_df = group_df.sort_values(
            by=['会計年度', 'sort_key', '項目名(日本語)']
        ).drop(columns=['sort_key'])
        
        # 列順序の統一 (数値と文字のカラムを並べる)
        desired_cols = [
            '企業名', '証券コード', '会計年度', 'カテゴリ', '項目名(日本語)', 
            '値(数値)', '値(文字)', '単位', '単体連結区分', '詳細文脈', '元ファイル', '項目名(英語)'
        ]
        # 実際に存在するカラムのみを選択
        cols = [c for c in desired_cols if c in out_df.columns]
        out_df = out_df[cols]

        # ファイル名生成
        safe_name = str(company).replace('/', '・').replace('\\', '￥')
        fy_start = target_years[0] if target_years else "ALL"
        fy_end = target_years[-1] if target_years else "ALL"
        
        filename = f"{safe_name}_{code}_{fy_start}_{fy_end}.csv"
        filepath = os.path.join(save_dir, filename)
        
        out_df.to_csv(filepath, index=False, encoding='utf-8-sig')
        print(f"  -> Saved: {filename} ({len(out_df)} rows)")
        count_files += 1

    print(f"\n[Complete] {count_files} files saved. Returning DataFrame.")
    return df


# =============================================================================
# コアクラス定義 (TaxonomyManager, XbrlExtractor)
# =============================================================================

class TaxonomyManager:
    def __init__(self, taxonomy_dir, map_cache_dir, base_cache_path=None):
        self.taxonomy_dir = taxonomy_dir
        self.ext_cache_dir = map_cache_dir
        self.base_cache_file = base_cache_path if base_cache_path else DEFAULT_BASE_CACHE_NAME
        self.base_labels = {}
        self.base_categories = {}
        self.is_base_loaded = False
        if not os.path.exists(self.ext_cache_dir): os.makedirs(self.ext_cache_dir)

    def load_base_taxonomy(self):
        if os.path.exists(self.base_cache_file):
            print(f"[System] 汎用タクソノミキャッシュをロード中...")
            self._load_base_json()
            print(f" -> 完了. 定義数: {len(self.base_labels)}")
        else:
            print(f"[System] 初回構築: 金融庁タクソノミ(Zip)を解析中...")
            t_start = time.time()
            zip_path = self._find_taxonomy_zip(self.taxonomy_dir)
            if not zip_path: raise FileNotFoundError(f"Taxonomy Zip not found in {self.taxonomy_dir}")
            self._build_base_from_zip(zip_path)
            self._apply_category_overrides()
            if len(self.base_labels) > 0:
                self._save_base_json()
                print(f" -> 解析完了 ({time.time()-t_start:.1f}s). キャッシュ保存済み。")
            else:
                print("【警告】ラベル定義が取得できませんでした。パスを確認してください。")
        self.is_base_loaded = True
        
    def _load_base_json(self):
        with open(self.base_cache_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            self.base_labels = data.get('labels', {})
            self.base_categories = data.get('categories', {})

    def _save_base_json(self):
        with open(self.base_cache_file, 'w', encoding='utf-8') as f:
            json.dump({'labels': self.base_labels, 'categories': self.base_categories}, f, ensure_ascii=False)
            
    def build_company_specific_caches(self, xbrl_zip_paths, target_companies, rebuild=False):
        print(f"[System] 企業別拡張タクソノミの準備 (Rebuild={rebuild})...")
        new_cnt, skip_cnt = 0, 0
        for zp in xbrl_zip_paths:
            try:
                base_name = os.path.basename(zp).split('.')[0]
                with zipfile.ZipFile(zp, 'r') as z:
                    xfiles = [f for f in z.namelist() if f.endswith('.xbrl') and 'PublicDoc' in f]
                    if not xfiles: continue
                    with z.open(xfiles[0]) as f:
                        soup = BeautifulSoup(f, 'lxml-xml')
                        fl = soup.find(re.compile(r'.*FilerNameInJapanese'))
                        cd = soup.find(re.compile(r'.*SecurityCode'))
                        ftxt = fl.text.strip() if fl else ""
                        ctxt = cd.text.strip() if cd else ""
                        
                        if not any(k.strip() in f"{ftxt} {ctxt}" for p in target_companies for k in p.split('|')):
                            continue
                        
                        s_code = ctxt if ctxt else base_name
                        c_path = os.path.join(self.ext_cache_dir, f"map_{s_code}.json")
                        
                        if not rebuild and os.path.exists(c_path):
                            skip_cnt += 1
                            continue
                        
                        ext_labels = {}
                        lfiles = [n for n in z.namelist() if '_lab' in n and n.endswith('.xml')]
                        for lf in lfiles:
                            with z.open(lf) as obj: self._parse_label_stream_with_arc(obj, ext_labels)
                        
                        with open(c_path, 'w', encoding='utf-8') as f:
                            json.dump(ext_labels, f, ensure_ascii=False)
                        new_cnt += 1
            except: continue
        print(f" -> 新規作成: {new_cnt}, 既存利用: {skip_cnt}")

    def get_combined_map(self, company_code):
        merged = self.base_labels.copy()
        ep = os.path.join(self.ext_cache_dir, f"map_{company_code}.json")
        if os.path.exists(ep):
            with open(ep, 'r', encoding='utf-8') as f: merged.update(json.load(f))
        return merged, self.base_categories

    def _find_taxonomy_zip(self, d):
        z = glob.glob(os.path.join(d, "*.zip"))
        return z[0] if z else None

    def _build_base_from_zip(self, zp):
        with zipfile.ZipFile(zp, 'r') as z:
            lfiles = [f for f in z.namelist() if 'taxonomy/' in f and (f.endswith('_lab.xml') or f.endswith('_gla.xml') or f.endswith('_lab-g-ja.xml'))]
            print(f"  -> ラベル定義ファイル: {len(lfiles)} 個")
            for i, f in enumerate(lfiles):
                if i%100==0: print(f"    Parsing labels {i}/{len(lfiles)}...", end='\r')
                with z.open(f) as obj: self._parse_label_stream_with_arc(obj, self.base_labels)
            
            pfiles = [f for f in z.namelist() if 'taxonomy/' in f and f.endswith('_pre.xml')]
            print(f"\n  -> 構造定義ファイル: {len(pfiles)} 個")
            for i, f in enumerate(pfiles):
                if i%100==0: print(f"    Parsing structures {i}/{len(pfiles)}...", end='\r')
                with z.open(f) as obj: self._parse_pre_stream(obj, self.base_categories)
            print("")

    def _parse_label_stream_with_arc(self, fobj, ldict):
        try:
            soup = BeautifulSoup(fobj, 'lxml-xml')
            t_lbls, t_locs = {}, {}
            for t in soup.find_all(re.compile(r'.*label$')):
                lng = t.get('xml:lang') or t.get('lang')
                if lng and lng.startswith('ja'):
                    lid = t.get('xlink:label') or t.get('label')
                    if lid and t.text: t_lbls[lid] = t.text.strip()
            for t in soup.find_all(re.compile(r'.*loc$')):
                href = t.get('xlink:href') or t.get('href')
                if href and '#' in href:
                    lid = t.get('xlink:label') or t.get('label')
                    if lid: t_locs[lid] = href.split('#')[1]
            for a in soup.find_all(re.compile(r'.*labelArc$')):
                frm, to = a.get('xlink:from') or a.get('from'), a.get('xlink:to') or a.get('to')
                if frm in t_locs and to in t_lbls:
                    self._reg_label(ldict, t_locs[frm], t_lbls[to])
            # Fallback
            for lid, txt in t_lbls.items():
                if lid.startswith('label_'): self._reg_label(ldict, lid.replace('label_', ''), txt)
        except: pass

    def _reg_label(self, ldict, raw, txt):
        ldict[raw] = txt
        if '_' in raw:
            simple = raw.split('_')[-1]
            if len(simple)>2 and not simple.isdigit() and simple not in ldict: ldict[simple] = txt
            for p in ['jppfs_cor_', 'jpcrp_cor_', 'jpcrp', 'jpsps_cor_', 'jpigp_cor_']:
                if raw.startswith(p):
                    cl = raw.replace(p, '')
                    if '_' in cl: cl = cl.split('_')[-1]
                    ldict[cl] = txt

    def _parse_pre_stream(self, fobj, cdict):
        try:
            soup = BeautifulSoup(fobj, 'lxml-xml')
            for pl in soup.find_all(re.compile(r'.*presentationLink$')):
                cat = self._judge_role(pl.get('xlink:role', ''))
                if not cat: continue
                for loc in pl.find_all(re.compile(r'.*loc$')):
                    href = loc.get('xlink:href')
                    if href and '#' in href: self._upd_cat(cdict, href.split('#')[1], cat)
        except: pass

    def _upd_cat(self, cdict, tag, new_cat):
        prio = {'BS':1, 'PL':1, 'CF':1, 'SS':2, 'Notes':3, 'Others':9}
        curr = cdict.get(tag)
        if prio.get(new_cat, 99) <= prio.get(curr, 99):
            cdict[tag] = new_cat
            if '_' in tag: cdict[tag.split('_')[-1]] = new_cat

    def _judge_role(self, u):
        u = u.lower()
        if 'balancesheet' in u: return 'BS'
        if 'statementofincome' in u or 'profitandloss' in u or 'comprehensiveincome' in u: return 'PL'
        if 'cashflow' in u: return 'CF'
        if 'changesinequity' in u: return 'SS'
        if 'notes' in u: return 'Notes'
        return None

    def _apply_category_overrides(self):
        pl_re = re.compile(r'(ProfitLoss|NetIncome|OperatingIncome|NetSales|Revenue|CostOfSales|GrossProfit)$', re.IGNORECASE)
        bs_re = re.compile(r'(Assets|Liabilities|NetAssets|CashAndDeposits|RetainedEarnings)$', re.IGNORECASE)
        for t, c in self.base_categories.items():
            if pl_re.search(t) and c!='PL': self.base_categories[t]='PL'
            elif bs_re.search(t) and c!='BS': self.base_categories[t]='BS'

class XbrlExtractor:
    def __init__(self, tm):
        self.tm = tm

    def extract(self, zip_files, target_patterns, target_years):
        data = []
        print(f"[System] {len(zip_files)} ファイルの抽出処理を開始...")
        for zp in zip_files:
            try:
                rows = self._parse(zp, target_patterns, target_years)
                if rows:
                    data.extend(rows)
                    print(f"  [Hit] {rows[0]['企業名']} ({len(rows)} rows) - {os.path.basename(zp)}")
            except: continue
        return data

    def _parse(self, zp, pats, years):
        fname = os.path.basename(zp)
        with zipfile.ZipFile(zp, 'r') as z:
            xfiles = [f for f in z.namelist() if f.endswith('.xbrl') and 'PublicDoc' in f]
            if not xfiles: return []
            with z.open(xfiles[0]) as f:
                soup = BeautifulSoup(f, 'lxml-xml')
            
            fl = soup.find(re.compile(r'.*FilerNameInJapanese'))
            cd = soup.find(re.compile(r'.*SecurityCode'))
            ftxt = fl.text.strip() if fl else "Unknown"
            ctxt = cd.text.strip() if cd else ""
            
            if not any(k.strip() in f"{ftxt} {ctxt}" for p in pats for k in p.split('|')): return []
            
            lmap, cmap = self.tm.get_combined_map(ctxt)
            
            ctxs = {}
            for c in soup.find_all(re.compile(r'.*context$')):
                cid = c.get('id')
                is_con = True
                mems = []
                for m in c.find_all(re.compile(r'.*explicitMember$')):
                    v = m.text.strip().split(':')[-1]
                    if 'NonConsolidated' in v: is_con = False
                    if v not in ['ConsolidatedMember', 'NonConsolidatedMember']: mems.append(v)
                
                fy, pdate = "-", None
                p = c.find(re.compile(r'.*period$'))
                if p:
                    i = p.find(re.compile(r'.*instant$'))
                    e = p.find(re.compile(r'.*endDate$'))
                    dstr = i.text if i else (e.text if e else "")
                    if dstr:
                        try: pdate = datetime.strptime(dstr, '%Y-%m-%d')
                        except: pass
                if pdate: fy = f"FY{pdate.year - 1 if pdate.month <= 3 else pdate.year}"
                ctxs[cid] = {'type': "連結" if is_con else "単体", 'fy': fy, 'det': ", ".join(mems) if mems else "-"}

            res = []
            for t in soup.find_all():
                cref = t.get('contextRef')
                if not cref or cref not in ctxs: continue
                info = ctxs[cref]
                if years and info['fy'] not in years: continue
                
                # --- ★修正ポイント: 値の取得ロジック（数値か文字かで分岐） ---
                val_str = t.text.strip()
                if not val_str: continue # 空ならスキップ

                is_numeric = re.match(r'^-?\d+(\.\d+)?$', val_str)
                
                val_num = None
                val_text = None

                if is_numeric:
                    val_num = float(val_str)
                else:
                    # 自然言語（TextBlock等）の処理
                    # HTMLタグが含まれる、またはTextBlockタグの場合はタグを除去してテキスト化
                    if "TextBlock" in t.name or "<" in val_str:
                        # タグを除去して純粋なテキストのみ取得
                        val_text = BeautifulSoup(val_str, "lxml").get_text(" ", strip=True)
                    else:
                        val_text = val_str
                
                # -----------------------------------------------------------
                
                tname = t.name
                lbl = "-"
                if tname in lmap: lbl = lmap[tname]
                elif t.prefix and f"{t.prefix}_{tname}" in lmap: lbl = lmap[f"{t.prefix}_{tname}"]
                
                cat = cmap.get(tname)
                if not cat and t.prefix: cat = cmap.get(f"{t.prefix}_{tname}")
                if not cat: cat = self._guess(lbl)
                
                res.append({
                    '企業名': ftxt, '証券コード': ctxt, '会計年度': info['fy'], 'カテゴリ': cat,
                    '項目名(日本語)': lbl, 
                    '値(数値)': val_num,   # 数値カラム
                    '値(文字)': val_text,  # 文字カラム
                    '単位': t.get('unitRef'),
                    '単体連結区分': info['type'], '詳細文脈': info['det'], '元ファイル': fname,
                    '項目名(英語)': f"{t.prefix}:{t.name}" if t.prefix else t.name
                })
        return res

    def _guess(self, l):
        if l == "-": return "Others"
        if "キャッシュ・フロー" in l: return "CF"
        if any(k in l for k in ["資産", "負債", "純資産", "資本", "引当金", "未払", "未収"]): 
            if not any(k in l for k in ["益", "損", "費"]): return "BS"
        if any(k in l for k in ["売上", "収益", "費用", "利益", "損失"]): return "PL"
        return "Others"