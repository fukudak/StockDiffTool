import streamlit as st
import pandas as pd
import numpy as np
import logging
from typing import Dict, List, Tuple, Optional, Any
import hashlib
from dataclasses import dataclass
import math

# ログ設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# =============================================================================
# 設定・定数定義
# =============================================================================

class Config:
    """アプリケーション設定クラス"""
    
    # 基本設定
    FIXED_KEY_COLUMNS = ["品番", "サイズ枠", "カラーコード", "サイズ名", "ＪＡＮコード"]
    STOCK_COLUMN = "在庫数"
    
    # ファイル制限
    MAX_FILE_SIZE_MB = 50
    MAX_DATA_ROWS = 100000
    
    # ページング設定
    ITEMS_PER_PAGE = 100
    
    # UI設定
    DANGEROUS_CHARS = ['<', '>', ':', '"', '|', '?', '*', '\\', '/']
    TYPE_DISPLAY_MAP = {
        'added': '➕ 追加',
        'deleted': '➖ 削除', 
        'modified': '🔄 在庫変更',
        'unchanged': '✅ 変更なし'
    }
    TYPE_EXPORT_MAP = {
        'added': '追加',
        'deleted': '削除',
        'modified': '在庫変更', 
        'unchanged': '変更なし'
    }
    ENCODING_MAP = {
        "UTF-8 (BOM付き)": "utf-8-sig",
        "Shift_JIS": "shift_jis"
    }

@dataclass
class ComparisonResult:
    """比較結果データクラス"""
    type: str
    data: Dict[str, Any]
    stock1: float
    stock2: float
    stock1_display: str
    stock2_display: str
    stock_change: float
    key: str

@dataclass
class Summary:
    """サマリーデータクラス"""
    added: int = 0
    deleted: int = 0
    modified: int = 0
    unchanged: int = 0
    
    @property
    def total_items(self) -> int:
        return self.added + self.deleted + self.modified + self.unchanged

# =============================================================================
# ユーティリティクラス
# =============================================================================

class SessionManager:
    """セッション状態管理クラス"""
    
    # セッションキー
    COMPARISON_COMPLETED = "comparison_completed"
    ALL_ITEMS = "all_items"
    SUMMARY = "summary"
    ORIGINAL_COLUMNS = "original_columns"
    FILE1_NAME = "file1_name"
    FILE1_SHEET = "file1_sheet"
    FILE2_NAME = "file2_name"
    FILE2_SHEET = "file2_sheet"
    
    @classmethod
    def initialize(cls):
        """セッション状態を初期化"""
        if cls.COMPARISON_COMPLETED not in st.session_state:
            st.session_state[cls.COMPARISON_COMPLETED] = False
    
    @classmethod
    def clear(cls):
        """セッションデータをクリア"""
        keys_to_clear = [
            cls.COMPARISON_COMPLETED, cls.ALL_ITEMS, cls.SUMMARY, cls.ORIGINAL_COLUMNS,
            cls.FILE1_NAME, cls.FILE1_SHEET, cls.FILE2_NAME, cls.FILE2_SHEET
        ]
        for key in keys_to_clear:
            if key in st.session_state:
                del st.session_state[key]
    
    @classmethod
    def save_comparison_result(cls, all_items: List[ComparisonResult], summary: Summary,
                             original_columns: List[str], file1_name: str, file1_sheet: str,
                             file2_name: str, file2_sheet: str):
        """比較結果をセッションに保存"""
        st.session_state.update({
            cls.ALL_ITEMS: all_items,
            cls.SUMMARY: summary,
            cls.ORIGINAL_COLUMNS: original_columns,
            cls.COMPARISON_COMPLETED: True,
            cls.FILE1_NAME: file1_name,
            cls.FILE1_SHEET: file1_sheet,
            cls.FILE2_NAME: file2_name,
            cls.FILE2_SHEET: file2_sheet
        })

class FileUtils:
    """ファイル関連ユーティリティ"""
    
    @staticmethod
    def validate_file(uploaded_file) -> Tuple[bool, str]:
        """ファイルの基本検証"""
        # ファイルサイズチェック
        file_size_mb = uploaded_file.size / (1024 * 1024)
        if file_size_mb > Config.MAX_FILE_SIZE_MB:
            return False, f"ファイルサイズが制限を超えています（{file_size_mb:.1f}MB > {Config.MAX_FILE_SIZE_MB}MB）"
        
        # ファイル名の基本チェック
        if any(char in uploaded_file.name for char in Config.DANGEROUS_CHARS):
            return False, "ファイル名に不正な文字が含まれています"
        
        return True, "OK"
    
    @staticmethod
    def load_excel_file(uploaded_file, file_identifier: str) -> Tuple[Optional[pd.DataFrame], str, str]:
        """Excelファイルの読み込み"""
        try:
            # 基本的なファイル検証
            is_valid, error_msg = FileUtils.validate_file(uploaded_file)
            if not is_valid:
                st.error(f"❌ {error_msg}")
                return None, "", ""
            
            excel_file = pd.ExcelFile(uploaded_file)
            if not excel_file.sheet_names:
                st.warning("シートが見つかりません")
                return None, "", ""
            
            # シート選択
            sheet_name = st.selectbox(
                "シートを選択",
                excel_file.sheet_names,
                key=f"sheet_{file_identifier}_{hashlib.md5(uploaded_file.name.encode()).hexdigest()[:8]}",
                help="分析対象のシートを選択してください"
            )
            
            if sheet_name:
                data = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str).fillna("")
                
                # データサイズ制限
                if len(data) > Config.MAX_DATA_ROWS:
                    st.warning(f"⚠️ データが大きすぎます。最初の{Config.MAX_DATA_ROWS:,}行のみ読み込まれました。")
                    data = data.head(Config.MAX_DATA_ROWS)
                
                st.success(f"✅ 「{sheet_name}」から {len(data):,} 行、{len(data.columns)} 列を読み込みました")
                
                # データプレビュー
                with st.expander("📊 データプレビュー"):
                    st.dataframe(data.head(10), use_container_width=True)
                
                return data, uploaded_file.name, sheet_name
            
            return None, uploaded_file.name, ""
                
        except Exception as e:
            st.error(f"ファイル読み込みエラー: {str(e)}")
            logger.error(f"File loading error: {e}")
            return None, "", ""

class StockUtils:
    """在庫関連ユーティリティ"""
    
    @staticmethod
    def get_stock_value(row: Optional[Dict], stock_column: str, default: float = 0.0) -> Tuple[float, str]:
        """在庫数を安全に取得し、元の値も返す"""
        if not row:
            return default, ""
        
        value = row.get(stock_column, default)
        original_value = str(value).strip()
        
        # ●の場合はそのまま返す
        if original_value == "●":
            return 0.0, "●"
        
        if pd.isna(value) or value == "":
            return default, ""
        
        try:
            if isinstance(value, str):
                cleaned_value = ''.join(filter(lambda x: x.isdigit() or x == '.', value))
                numeric_value = float(cleaned_value) if cleaned_value else default
                return numeric_value, original_value
            return float(value), original_value
        except (ValueError, TypeError):
            return default, original_value
    
    @staticmethod
    def format_stock_display(item: ComparisonResult) -> Tuple[str, str, str]:
        """在庫表示値をフォーマット"""
        if item.type == 'added':
            stock1_display = ""
            stock2_display = item.stock2_display if item.stock2_display == "●" else f"{item.stock2:.0f}"
            stock_change_display = ""
        elif item.type == 'deleted':
            stock1_display = item.stock1_display if item.stock1_display == "●" else f"{item.stock1:.0f}"
            stock2_display = ""
            stock_change_display = ""
        else:  # modified, unchanged
            stock1_display = item.stock1_display if item.stock1_display == "●" else f"{item.stock1:.0f}"
            stock2_display = item.stock2_display if item.stock2_display == "●" else f"{item.stock2:.0f}"
            stock_change_display = "" if (item.stock1_display == "●" or item.stock2_display == "●") else f"{item.stock_change:+.0f}"
        
        return stock1_display, stock2_display, stock_change_display

class PaginationUtils:
    """ページング関連ユーティリティ"""
    
    @staticmethod
    def get_page_info(total_items: int, items_per_page: int = Config.ITEMS_PER_PAGE) -> Tuple[int, int]:
        """総ページ数と最大ページ番号を取得"""
        total_pages = math.ceil(total_items / items_per_page) if total_items > 0 else 1
        max_page = max(1, total_pages)
        return total_pages, max_page
    
    @staticmethod
    def get_page_items(items: List[Any], page: int, items_per_page: int = Config.ITEMS_PER_PAGE) -> List[Any]:
        """指定ページのアイテムを取得"""
        start_idx = (page - 1) * items_per_page
        end_idx = start_idx + items_per_page
        return items[start_idx:end_idx]
    
    @staticmethod
    def get_page_range_info(total_items: int, page: int, items_per_page: int = Config.ITEMS_PER_PAGE) -> Tuple[int, int]:
        """現在ページの表示範囲を取得"""
        start_item = (page - 1) * items_per_page + 1
        end_item = min(page * items_per_page, total_items)
        return start_item, end_item

# =============================================================================
# コアビジネスロジック
# =============================================================================

class InventoryComparator:
    """在庫比較を行うクラス"""
    
    def __init__(self):
        self.key_columns = Config.FIXED_KEY_COLUMNS
        self.stock_column = Config.STOCK_COLUMN
    
    def validate_data(self, data1: pd.DataFrame, data2: pd.DataFrame) -> Tuple[bool, str]:
        """データの妥当性を検証"""
        if data1.empty or data2.empty:
            return False, "空のデータが含まれています"
        
        if set(data1.columns) != set(data2.columns):
            return False, "列構成が異なります"
        
        missing_keys = [key for key in self.key_columns if key not in data1.columns]
        if missing_keys:
            return False, f"必須キー列が不足: {missing_keys}"
        
        if self.stock_column not in data1.columns:
            return False, f"在庫列 '{self.stock_column}' が見つかりません"
        
        return True, "OK"
    
    def _generate_row_key(self, row: pd.Series) -> str:
        """行の一意キーを生成"""
        return "|".join([str(row.get(col, "")).strip() for col in self.key_columns])
    
    def _create_data_map(self, data: pd.DataFrame) -> Dict[str, Dict]:
        """データをマップに変換"""
        return {
            self._generate_row_key(row): row.to_dict() 
            for _, row in data.iterrows() 
            if self._generate_row_key(row)
        }
    
    def _determine_item_type(self, row1: Optional[Dict], row2: Optional[Dict], 
                           stock1_display: str, stock2_display: str) -> Tuple[str, Dict]:
        """アイテムタイプを判定"""
        if row1 and row2:
            item_type = 'modified' if stock1_display != stock2_display else 'unchanged'
            item_data = row2
        elif row2 and not row1:
            item_type = 'added'
            item_data = row2
        elif row1 and not row2:
            item_type = 'deleted'
            item_data = row1
        else:
            raise ValueError("Invalid row combination")
        
        return item_type, item_data
    
    def compare_inventories(self, data1: pd.DataFrame, data2: pd.DataFrame) -> Tuple[List[ComparisonResult], Summary]:
        """在庫比較の実行"""
        # データをマップに変換
        map1 = self._create_data_map(data1)
        map2 = self._create_data_map(data2)
        
        all_keys = sorted(set(map1.keys()) | set(map2.keys()))
        
        all_items = []
        summary = Summary()
        
        for key in all_keys:
            row1 = map1.get(key)
            row2 = map2.get(key)
            
            stock1_numeric, stock1_display = StockUtils.get_stock_value(row1, self.stock_column)
            stock2_numeric, stock2_display = StockUtils.get_stock_value(row2, self.stock_column)
            stock_change = stock2_numeric - stock1_numeric
            
            try:
                item_type, item_data = self._determine_item_type(
                    row1, row2, stock1_display, stock2_display
                )
                
                # サマリー更新
                if item_type == 'added':
                    summary.added += 1
                elif item_type == 'deleted':
                    summary.deleted += 1
                elif item_type == 'modified':
                    summary.modified += 1
                else:
                    summary.unchanged += 1
                
                all_items.append(ComparisonResult(
                    type=item_type,
                    data=item_data,
                    stock1=stock1_numeric,
                    stock2=stock2_numeric,
                    stock1_display=stock1_display,
                    stock2_display=stock2_display,
                    stock_change=stock_change,
                    key=key
                ))
                
            except ValueError:
                continue
        
        return all_items, summary

# =============================================================================
# UI コンポーネント
# =============================================================================

class UIRenderer:
    """UI描画クラス"""
    
    @staticmethod
    def render_header(comparison_completed: bool):
        """ヘッダー表示"""
        st.markdown("""
        <div style="background: linear-gradient(90deg, #4b6cb7 0%, #182848 100%); padding: 0.8rem; border-radius: 8px; margin-bottom: 1rem;">
            <h2 style="color: white; text-align: center; margin: 0;">📊 在庫差分比較ツール</h2>
            <p style="color: #E0E0E0; text-align: center; margin: 0.3rem 0 0 0; font-size: 0.9rem;">Excelファイルの在庫データを比較し、差分を可視化するツールです。</p>
        </div>
        """, unsafe_allow_html=True)
        
        if comparison_completed and all(
            key in st.session_state for key in [
                SessionManager.FILE1_NAME, SessionManager.FILE1_SHEET,
                SessionManager.FILE2_NAME, SessionManager.FILE2_SHEET
            ]
        ):
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**📁 ファイル1:** {st.session_state[SessionManager.FILE1_NAME]}")
                st.markdown(f"**📄 シート:** {st.session_state[SessionManager.FILE1_SHEET]}")
            with col2:
                st.markdown(f"**📁 ファイル2:** {st.session_state[SessionManager.FILE2_NAME]}")
                st.markdown(f"**📄 シート:** {st.session_state[SessionManager.FILE2_SHEET]}")
    
    @staticmethod
    def render_sidebar(comparison_completed: bool) -> None:
        """サイドバー表示"""
        with st.sidebar:
            st.markdown("### ⚙️ 設定情報")
            st.markdown("**比較項目:**")
            st.code(Config.STOCK_COLUMN, language="text")
            st.markdown("**比較キー:**")
            st.code("\n".join(Config.FIXED_KEY_COLUMNS), language="text")
            
            st.markdown("---")
            st.markdown("### 📥 ダウンロード設定")
            encoding = st.selectbox("文字コードを選択", ["UTF-8 (BOM付き)", "Shift_JIS"])
            
            # ダウンロードボタン
            if comparison_completed and SessionManager.ALL_ITEMS in st.session_state:
                csv_data = CSVExporter.create_csv(
                    st.session_state[SessionManager.ALL_ITEMS], 
                    st.session_state[SessionManager.ORIGINAL_COLUMNS], 
                    encoding
                )
                st.download_button(
                    label="📥 CSVダウンロード",
                    data=csv_data,
                    file_name=f"inventory_comparison_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            else:
                st.button("📥 CSVダウンロード", disabled=True, use_container_width=True)
    
    @staticmethod
    def render_file_upload_section() -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame], str, str, str, str]:
        """ファイルアップロードセクションの表示"""
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📁 ファイル1（比較元）")
            file1 = st.file_uploader("Excelファイル1", type=['xlsx', 'xls'], key="file1")
            data1, file1_name, file1_sheet = FileUtils.load_excel_file(file1, "file1") if file1 else (None, "", "")
        
        with col2:
            st.subheader("📁 ファイル2（比較先）")
            file2 = st.file_uploader("Excelファイル2", type=['xlsx', 'xls'], key="file2")
            data2, file2_name, file2_sheet = FileUtils.load_excel_file(file2, "file2") if file2 else (None, "", "")
        
        return data1, data2, file1_name, file1_sheet, file2_name, file2_sheet
    
    @staticmethod
    def render_pagination_controls(total_items: int, current_page: int, tab_key: str) -> int:
        """ページング制御の表示"""
        if total_items <= Config.ITEMS_PER_PAGE:
            return current_page
        
        total_pages, max_page = PaginationUtils.get_page_info(total_items)
        start_item, end_item = PaginationUtils.get_page_range_info(total_items, current_page)
        
        # ページング情報の表示
        col1, col2, col3 = st.columns([2, 1, 2])
        
        with col1:
            st.markdown(f"**表示範囲:** {start_item:,} - {end_item:,} 件 / 全 {total_items:,} 件")
        
        with col2:
            st.markdown(f"**ページ:** {current_page} / {total_pages}")
        
        with col3:
            # ページ選択
            new_page = st.selectbox(
                "ページ選択",
                range(1, total_pages + 1),
                index=current_page - 1,
                key=f"page_select_{tab_key}",
                label_visibility="collapsed"
            )
        
        # ナビゲーションボタン
        nav_col1, nav_col2, nav_col3, nav_col4, nav_col5 = st.columns(5)
        
        with nav_col1:
            if st.button("⏮️ 最初", key=f"first_{tab_key}", disabled=(current_page == 1)):
                new_page = 1
        
        with nav_col2:
            if st.button("◀️ 前", key=f"prev_{tab_key}", disabled=(current_page == 1)):
                new_page = max(1, current_page - 1)
        
        with nav_col4:
            if st.button("▶️ 次", key=f"next_{tab_key}", disabled=(current_page == total_pages)):
                new_page = min(total_pages, current_page + 1)
        
        with nav_col5:
            if st.button("⏭️ 最後", key=f"last_{tab_key}", disabled=(current_page == total_pages)):
                new_page = total_pages
        
        return new_page
    
    @staticmethod
    def render_results(all_items: List[ComparisonResult], original_columns: List[str], summary: Summary):
        """結果表示"""
        if not all_items:
            st.info("🎉 比較対象のアイテムが見つかりませんでした")
            return
        
        st.markdown("## 📋 比較結果")
        
        # タブ設定
        tab_configs = [
            (f"🔍 全てのアイテム ({summary.total_items:,}件)", "all"),
            (f"➕ 追加 ({summary.added:,}件)", "added"),
            (f"➖ 削除 ({summary.deleted:,}件)", "deleted"),
            (f"🔄 在庫変更 ({summary.modified:,}件)", "modified"),
            (f"✅ 変更なし ({summary.unchanged:,}件)", "unchanged")
        ]
        
        tabs = st.tabs([config[0] for config in tab_configs])
        
        for tab, (_, filter_type) in zip(tabs, tab_configs):
            with tab:
                if filter_type == "all":
                    filtered_items = all_items
                else:
                    filtered_items = [item for item in all_items if item.type == filter_type]
                
                if not filtered_items:
                    st.info("該当するアイテムはありません")
                    continue
                
                # ページング状態の管理
                page_key = f"page_{filter_type}"
                if page_key not in st.session_state:
                    st.session_state[page_key] = 1
                
                # ページング制御
                current_page = UIRenderer.render_pagination_controls(
                    len(filtered_items), 
                    st.session_state[page_key], 
                    filter_type
                )
                
                # ページが変更された場合は更新
                if current_page != st.session_state[page_key]:
                    st.session_state[page_key] = current_page
                    st.rerun()
                
                # 現在ページのアイテムを取得
                page_items = PaginationUtils.get_page_items(filtered_items, current_page)
                
                if not page_items:
                    st.info("このページには表示するアイテムがありません")
                    continue
                
                # 表示データの準備
                display_data = []
                for item in page_items:
                    stock1_display, stock2_display, stock_change_display = StockUtils.format_stock_display(item)
                    
                    row = {
                        '変更タイプ': Config.TYPE_DISPLAY_MAP[item.type],
                        '在庫(元)': stock1_display,
                        '在庫(先)': stock2_display,
                        '在庫変化': stock_change_display
                    }
                    
                    # 元の列データを追加
                    for col in original_columns:
                        row[col] = item.data.get(col, '')
                    
                    display_data.append(row)
                
                df = pd.DataFrame(display_data)
                st.dataframe(
                    df, 
                    use_container_width=True, 
                    hide_index=True, 
                    height=min(600, len(df) * 35 + 38)
                )
                
                # 現在の表示情報
                start_item, end_item = PaginationUtils.get_page_range_info(len(filtered_items), current_page)
                st.caption(f"現在のページ: {start_item:,} - {end_item:,} 件 / 全 {len(filtered_items):,} 件")

class CSVExporter:
    """CSV出力クラス"""
    
    @staticmethod
    def create_csv(all_items: List[ComparisonResult], original_columns: List[str], encoding_choice: str) -> bytes:
        """CSV出力の生成"""
        if not all_items:
            return b""
        
        export_data = []
        for item in all_items:
            stock1_display, stock2_display, stock_change_display = StockUtils.format_stock_display(item)
            
            row = {
                '変更タイプ': Config.TYPE_EXPORT_MAP[item.type],
                'ファイル1在庫': stock1_display,
                'ファイル2在庫': stock2_display,
                '在庫変化': stock_change_display,
                '比較キー': item.key,
            }
            
            # 元の列データを追加
            for col in original_columns:
                row[col] = item.data.get(col, '')
            
            export_data.append(row)
        
        df = pd.DataFrame(export_data)
        encoding = Config.ENCODING_MAP.get(encoding_choice, "utf-8-sig")
        return df.to_csv(index=False).encode(encoding)

# =============================================================================
# メインアプリケーション
# =============================================================================

class InventoryComparisonApp:
    """メインアプリケーションクラス"""
    
    def __init__(self):
        self.comparator = InventoryComparator()
    
    def _setup_page(self):
        """ページ設定"""
        st.set_page_config(
            page_title="在庫差分比較ツール v6.1",
            page_icon="📊",
            layout="wide"
        )
        
        # カスタムCSS
        st.markdown("""
        <style>
        .stButton>button { 
            border-radius: 8px; 
            font-weight: bold; 
        }
        .main > div {
            padding-top: 1rem;
        }
        .stMetric {
            background-color: #f0f2f6;
            padding: 0.5rem;
            border-radius: 4px;
        }
        </style>
        """, unsafe_allow_html=True)
    
    def _handle_comparison_execution(self, data1: pd.DataFrame, data2: pd.DataFrame, 
                                   file1_name: str, file1_sheet: str, 
                                   file2_name: str, file2_sheet: str):
        """比較実行処理"""
        is_valid, error_msg = self.comparator.validate_data(data1, data2)
        
        if not is_valid:
            st.error(f"❌ データ検証エラー: {error_msg}")
            return
        
        if st.button("🔍 比較実行", type="primary", use_container_width=True):
            with st.spinner("比較処理を実行中..."):
                try:
                    all_items, summary = self.comparator.compare_inventories(data1, data2)
                    
                    # 結果を保存
                    SessionManager.save_comparison_result(
                        all_items, summary, data1.columns.tolist(),
                        file1_name, file1_sheet, file2_name, file2_sheet
                    )
                    
                    st.success("✅ 比較処理が完了しました！")
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ 比較処理中にエラーが発生しました: {str(e)}")
                    logger.error(f"Comparison error: {e}")
    
    def run(self):
        """アプリケーション実行"""
        # 初期設定
        self._setup_page()
        SessionManager.initialize()
        
        # UI描画
        comparison_completed = st.session_state[SessionManager.COMPARISON_COMPLETED]
        UIRenderer.render_header(comparison_completed)
        
        # サイドバー
        UIRenderer.render_sidebar(comparison_completed)
        
        # メインコンテンツ
        if not comparison_completed:
            # ファイルアップロードと比較実行
            data1, data2, file1_name, file1_sheet, file2_name, file2_sheet = UIRenderer.render_file_upload_section()
            
            if data1 is not None and data2 is not None:
                self._handle_comparison_execution(data1, data2, file1_name, file1_sheet, file2_name, file2_sheet)
        else:
            # クリアボタン
            if st.button("🗑️ クリア", type="secondary", use_container_width=True):
                SessionManager.clear()
                st.rerun()
            
            # 結果表示
            UIRenderer.render_results(
                st.session_state[SessionManager.ALL_ITEMS], 
                st.session_state[SessionManager.ORIGINAL_COLUMNS], 
                st.session_state[SessionManager.SUMMARY]
            )

def main():
    """メイン関数"""
    app = InventoryComparisonApp()
    app.run()

if __name__ == "__main__":
    main()
