import pandas as pd
import re

class FinancialAnalyzer:
    """
    EDINET抽出データ(DataFrame)を受け取り、財務分析を行うクラス。
    
    特徴:
    - 企業ごとの会計基準(IFRS/US GAAP/JGAAP)や業種(銀行)の差異を吸収
    - 銀行特有のBS項目(預金等)にも対応
    - 存在しないデータはNoneとして扱う
    """
    
    def __init__(self, df_input):
        self.raw_df = df_input.copy()
        
        # =========================================================
        # 1. PL (損益計算書) マッピング
        # =========================================================
        # ★修正: 変数名を self.pl_metrics に統一 (AttributeError回避)
        self.pl_metrics = {
            '売上高': {
                'tags': [
                    '^NetSalesSummary', '^RevenueIFRSSummary', '^OperatingRevenuesIFRSSummary',
                    '^Revenue', '^OperatingRevenue', '^SalesRevenuesIFRS', '^NetSales',
                    '^OrdinaryIncome'
                ]
            },
            '営業利益': {
                'tags': [
                    '^BusinessProfit', '^OperatingProfitLossIFRS', 
                    '^OperatingIncome', '^OperatingProfit',
                    '^OrdinaryProfit'
                ]
            },
            '当期純利益': {
                'tags': [
                    '^ProfitLossAttributableToOwnersOfParent', 
                    '^NetIncome', '^ProfitLoss', '^CurrentNetIncome'
                ]
            }
        }

        # =========================================================
        # 2. BS (貸借対照表) マッピング
        # =========================================================
        # ★修正: 変数名を self.bs_metrics に統一
        self.bs_metrics = {
            # --- 資産 ---
            '総資産': {'tags': ['^TotalAssets', '^Assets']},
            '流動資産': {'tags': ['^CurrentAssets']},
            
            # --- 負債 ---
            '流動負債': {'tags': ['^CurrentLiabilities']},
            
            # --- 純資産 ---
            '自己資本': { 
                # ROE計算の分母
                'tags': [
                    '^EquityAttributableToOwnersOfParent',
                    '^NetAssets', '^TotalNetAssets', '^ShareholdersEquity'
                ]
            },
            '純資産': { # 参考
                 'tags': ['^TotalNetAssets', '^NetAssets', '^TotalEquity', '^Equity']
            },
            
            # --- 銀行特有項目 ---
            '預金': {'tags': ['^DepositsLiabilitiesBNK']},
            '貸出金': {'tags': ['^LoansAndBillsDiscountedAssetsBNK']},
            '有価証券': {'tags': ['^SecuritiesAssetsBNK']},
            '現金預け金': {'tags': ['^CashAndDueFromBanksAssetsBNK']},
            '借用金': {'tags': ['^BorrowedMoneyLiabilitiesBNK']}
        }

        # =========================================================
        # 3. 企業別専用マッピング (Strict Mapping)
        # =========================================================
        self.company_specific_metrics = {
            # --- トヨタ自動車 (72030) ---
            '72030': {
                '売上高': ['^OperatingRevenuesIFRSKey', '^RevenuesUSGAAPSummary', '^SalesRevenuesIFRS'],
                '営業利益': ['^OperatingProfitLossIFRS', '^OperatingIncomeLoss'],
                '当期純利益': ['^ProfitLossAttributableToOwnersOfParentIFRS', '^NetIncomeLossAttributableToOwnersOfParentUSGAAP', '^ProfitLossAttributableToOwnersOfParent'],
                '総資産': ['^TotalAssetsIFRS', '^TotalAssetsUSGAAPSummary'],
                '自己資本': ['^EquityAttributableToOwnersOfParentIFRS', '^ShareholdersEquityUSGAAP', '^EquityIncludingPortionAttributableToNonControllingInterestUSGAAP'],
                '純資産': ['^TotalEquityIFRS', '^EquityIncludingPortionAttributableToNonControllingInterestUSGAAP']
            },

            # --- 三菱UFJ (83060) ---
            '83060': {
                '売上高': ['^OrdinaryIncomeSummary', '^OrdinaryIncome'],
                '営業利益': ['^OrdinaryIncomeLossSummary', '^OrdinaryProfit'],
                '当期純利益': ['^ProfitLossAttributableToOwnersOfParentSummary', '^NetIncome'],
                '総資産': ['^TotalAssetsSummary', '^TotalAssets'],
                '自己資本': ['^TotalNetAssetsSummary', '^NetAssets'],
                '純資産': ['^TotalNetAssetsSummary', '^NetAssets']
            },

            # --- 本田技研工業 (72670) ---
            '72670': {
                '売上高': ['^SalesRevenueIFRSSummary', '^RevenueIFRS', '^SalesRevenue'],
                '営業利益': ['^OperatingProfitIFRSSummary', '^OperatingProfitLossIFRS', '^OperatingProfit'],
                '当期純利益': ['^ProfitLossAttributableToOwnersOfParentIFRS'],
                '総資産': ['^TotalAssetsIFRS'],
                '自己資本': ['^EquityAttributableToOwnersOfParentIFRS'],
                '純資産': ['^TotalEquityIFRS']
            },

            # --- 味の素 (28020) ---
            '28020': {
                '売上高': ['^RevenueIFRS', '^SalesIFRSSummary', '^SalesToExternalCustomersIFRS'],
                '営業利益': ['^BusinessProfitIFRSSummary', '^BusinessProfit', '^OperatingProfit'],
                '当期純利益': ['^ProfitLossAttributableToOwnersOfParentIFRS'],
                '総資産': ['^TotalAssetsIFRS'],
                '自己資本': ['^EquityAttributableToOwnersOfParentIFRS', '^TotalEquityIFRSSummary'],
                '純資産': ['^TotalEquityIFRS']
            }
        }

    # =========================================================
    # メイン分析メソッド
    # =========================================================

    def analyze_pl(self):
        """PL分析を実行"""
        return self._analyze_generic(self.pl_metrics, 'PL')

    def analyze_bs(self):
        """BS分析を実行"""
        return self._analyze_generic(self.bs_metrics, 'BS')

    def calculate_efficiency_metrics(self, df_pl, df_bs):
        """
        PLとBSの結果を結合し、指標(ROE, ROA等)を計算する。
        """
        merge_keys = ['企業名', '証券コード', '会計年度', '決算区分']
        
        merged = pd.merge(
            df_pl, 
            df_bs, 
            on=merge_keys, 
            how='outer',
            suffixes=('', '_BS') 
        )
        
        def calc_row(row):
            # 値の取得
            net_income = row.get('当期純利益')
            sales = row.get('売上高')
            op_profit = row.get('営業利益')
            
            assets = row.get('総資産')
            equity = row.get('自己資本')
            if pd.isna(equity): equity = row.get('純資産')

            cur_assets = row.get('流動資産')
            cur_liabs = row.get('流動負債')
            
            # --- 指標計算 ---
            if equity and net_income and equity != 0:
                row['ROE'] = (net_income / equity * 100)
            else:
                row['ROE'] = None
            
            if assets and net_income and assets != 0:
                row['ROA'] = (net_income / assets * 100)
            else:
                row['ROA'] = None
                
            if assets and equity and assets != 0:
                row['自己資本比率'] = (equity / assets * 100)
            else:
                row['自己資本比率'] = None
            
            if cur_assets and cur_liabs and cur_liabs != 0:
                row['流動比率'] = (cur_assets / cur_liabs * 100)
            else:
                row['流動比率'] = None
                
            if sales and op_profit and sales != 0:
                row['事業利益率'] = (op_profit / sales * 100)
            
            if sales and net_income and sales != 0:
                row['純利益率'] = (net_income / sales * 100)

            return row

        result_df = merged.apply(calc_row, axis=1)
        
        # --- カラム順序の整理 ---
        base_cols = merge_keys
        
        # 定義辞書から動的にカラムリストを生成 (これで銀行項目なども漏れない)
        pl_cols = list(self.pl_metrics.keys())
        bs_cols = list(self.bs_metrics.keys())
        calc_cols = ['事業利益率', '純利益率', 'ROE', 'ROA', '自己資本比率', '流動比率']
        
        # メタデータ
        meta_cols = [c for c in result_df.columns if '_タグ' in c or '_ラベル' in c]
        
        final_cols = base_cols + pl_cols + bs_cols + calc_cols + sorted(meta_cols)
        final_cols = [c for c in final_cols if c in result_df.columns]
        
        return result_df[final_cols]

    # =========================================================
    # 内部ロジック
    # =========================================================
    
    def _analyze_generic(self, metrics_def, type_label):
        clean_df = self._preprocess_df(self.raw_df)
        summary_data = []
        
        for (code, year), group in clean_df.groupby(['証券コード', '会計年度']):
            company_name = group['企業名'].iloc[0]
            
            df_con = group[group['単体連結区分'] == '連結']
            df_non = group[group['単体連結区分'] == '単体']
            target_df = df_con if not df_con.empty else df_non
            status = '連結' if not df_con.empty else '単体'
            
            row_data = {
                '企業名': company_name, '証券コード': code, 
                '会計年度': year, '決算区分': status
            }
            
            str_code = str(code)
            current_metrics = metrics_def.copy()
            
            if str_code in self.company_specific_metrics:
                spec = self.company_specific_metrics[str_code]
                for k, v in spec.items():
                    if k in metrics_def:
                        if isinstance(v, list): current_metrics[k] = {'tags': v}
                        else: current_metrics[k] = v
            
            for metric_name, criteria in current_metrics.items():
                val, tag, label = self._find_best_match(target_df, criteria['tags'])
                row_data[metric_name] = val
                row_data[f'{metric_name}_タグ'] = tag
                row_data[f'{metric_name}_ラベル'] = label
            
            summary_data.append(row_data)
            
        return pd.DataFrame(summary_data)

    def _preprocess_df(self, df):
        df_clean = df[df['詳細文脈'] == '-'].copy()
        duplicate_keys = ['証券コード', '会計年度', 'タグ(要素名)', '単体連結区分']
        df_clean = df_clean.drop_duplicates(subset=duplicate_keys, keep='last')
        return df_clean

    def _find_best_match(self, df_target, tag_list):
        for tag_keyword in tag_list:
            matches = df_target[
                df_target['タグ(要素名)'].str.contains(tag_keyword, case=False, regex=True, na=False)
            ]
            if not matches.empty:
                row = matches.iloc[0]
                return row['値(数値)'], row['タグ(要素名)'], row['項目名(日本語)']
        return None, None, None