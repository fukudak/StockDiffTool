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
# 設定・データクラス
# =============================================================================

@dataclass
class AppConfig:
    """アプリケーション設定"""
    # 基本設定
    KEY_COLUMNS: List[str] = None
    STOCK_COLUMN: str = "在庫数"
    
    # ファイル制限
    MAX_FILE_SIZE_MB: int = 50
    MAX_DATA_ROWS: int = 100000
    
    # ページング設定
    ITEMS_PER_PAGE_MOBILE: int = 20
    ITEMS_PER_PAGE_TABLET: int = 50
    ITEMS_PER_PAGE_DESKTOP: int = 100
    
    # UI設定
    DANGEROUS_CHARS: List[str] = None
    TYPE_DISPLAY_MAP: Dict[str, str] = None
    TYPE_EXPORT_MAP: Dict[str, str] = None
    ENCODING_MAP: Dict[str, str] = None
    
    def __post_init__(self):
        if self.KEY_COLUMNS is None:
            self.KEY_COLUMNS = ["品番", "サイズ枠", "カラーコード", "サイズ名", "ＪＡＮコード"]
        
        if self.DANGEROUS_CHARS is None:
            self.DANGEROUS_CHARS = ['<', '>', ':', '"', '|', '?', '*', '\\', '/']
        
        if self.TYPE_DISPLAY_MAP is None:
            self.TYPE_DISPLAY_MAP = {
                'added': '➕ 追加',
                'deleted': '➖ 削除', 
                'modified': '🔄 在庫変更',
                'unchanged': '✅ 変更なし'
            }
        
        if self.TYPE_EXPORT_MAP is None:
            self.TYPE_EXPORT_MAP = {
                'added': '追加',
                'deleted': '削除',
                'modified': '在庫変更', 
                'unchanged': '変更なし'
            }
        
        if self.ENCODING_MAP is None:
            self.ENCODING_MAP = {
                "UTF-8 (BOM付き)": "utf-8-sig",
                "Shift_JIS": "shift_jis"
            }

@dataclass
class ComparisonItem:
    """比較結果アイテム"""
    type: str
    data: Dict[str, Any]
    stock1: float
    stock2: float
    stock1_display: str
    stock2_display: str
    stock_change: float
    key: str

@dataclass
class ComparisonSummary:
    """比較結果サマリー"""
    added: int = 0
    deleted: int = 0
    modified: int = 0
    unchanged: int = 0
    
    @property
    def total_items(self) -> int:
        return self.added + self.deleted + self.modified + self.unchanged

# グローバル設定インスタンス
CONFIG = AppConfig()

# =============================================================================
# セッション管理
# =============================================================================

class SessionState:
    """セッション状態管理"""
    
    # セッションキー定数
    COMPARISON_COMPLETED = "comparison_completed"
    ALL_ITEMS = "all_items"
    SUMMARY = "summary"
    ORIGINAL_COLUMNS = "original_columns"
    FILE1_NAME = "file1_name"
    FILE1_SHEET = "file1_sheet"
    FILE2_NAME = "file2_name"
    FILE2_SHEET = "file2_sheet"
    DEVICE_TYPE = "device_type"
    
    @classmethod
    def initialize(cls):
        """セッション状態初期化"""
        defaults = {
            cls.COMPARISON_COMPLETED: False,
            cls.DEVICE_TYPE: 'desktop'
        }
        
        for key, default_value in defaults.items():
            if key not in st.session_state:
                st.session_state[key] = default_value
    
    @classmethod
    def clear_comparison_data(cls):
        """比較データをクリア"""
        keys_to_clear = [
            cls.COMPARISON_COMPLETED, cls.ALL_ITEMS, cls.SUMMARY, cls.ORIGINAL_COLUMNS,
            cls.FILE1_NAME, cls.FILE1_SHEET, cls.FILE2_NAME, cls.FILE2_SHEET
        ]
        
        for key in keys_to_clear:
            if key in st.session_state:
                del st.session_state[key]
    
    @classmethod
    def save_comparison_result(cls, items: List[ComparisonItem], summary: ComparisonSummary,
                             columns: List[str], file1_name: str, file1_sheet: str,
                             file2_name: str, file2_sheet: str):
        """比較結果を保存"""
        st.session_state.update({
            cls.ALL_ITEMS: items,
            cls.SUMMARY: summary,
            cls.ORIGINAL_COLUMNS: columns,
            cls.COMPARISON_COMPLETED: True,
            cls.FILE1_NAME: file1_name,
            cls.FILE1_SHEET: file1_sheet,
            cls.FILE2_NAME: file2_name,
            cls.FILE2_SHEET: file2_sheet
        })

# =============================================================================
# ユーティリティ関数
# =============================================================================

def get_device_type() -> str:
    """現在のデバイスタイプを取得"""
    return st.session_state.get(SessionState.DEVICE_TYPE, 'desktop')

def get_items_per_page() -> int:
    """デバイスタイプに応じたページあたりアイテム数を取得"""
    device_type = get_device_type()
    
    if device_type == 'mobile':
        return CONFIG.ITEMS_PER_PAGE_MOBILE
    elif device_type == 'tablet':
        return CONFIG.ITEMS_PER_PAGE_TABLET
    else:
        return CONFIG.ITEMS_PER_PAGE_DESKTOP

def validate_file(uploaded_file) -> Tuple[bool, str]:
    """ファイル検証"""
    # ファイルサイズチェック
    file_size_mb = uploaded_file.size / (1024 * 1024)
    if file_size_mb > CONFIG.MAX_FILE_SIZE_MB:
        return False, f"ファイルサイズが制限を超えています（{file_size_mb:.1f}MB > {CONFIG.MAX_FILE_SIZE_MB}MB）"
    
    # ファイル名チェック
    if any(char in uploaded_file.name for char in CONFIG.DANGEROUS_CHARS):
        return False, "ファイル名に不正な文字が含まれています"
    
    return True, "OK"

def parse_stock_value(value: Any) -> Tuple[float, str]:
    """在庫値を解析"""
    if pd.isna(value) or value == "":
        return 0.0, ""
    
    original_value = str(value).strip()
    
    # ●の場合はそのまま返す
    if original_value == "●":
        return 0.0, "●"
    
    try:
        if isinstance(value, str):
            cleaned_value = ''.join(filter(lambda x: x.isdigit() or x == '.', value))
            numeric_value = float(cleaned_value) if cleaned_value else 0.0
            return numeric_value, original_value
        return float(value), original_value
    except (ValueError, TypeError):
        return 0.0, original_value

def format_stock_display(item: ComparisonItem) -> Tuple[str, str, str]:
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

def get_page_info(total_items: int, items_per_page: int = None) -> Tuple[int, int, int, int]:
    """ページング情報を取得"""
    if items_per_page is None:
        items_per_page = get_items_per_page()
    
    total_pages = math.ceil(total_items / items_per_page) if total_items > 0 else 1
    max_page = max(1, total_pages)
    
    return total_pages, max_page, items_per_page, total_items

def get_page_items(items: List[Any], page: int, items_per_page: int = None) -> List[Any]:
    """指定ページのアイテムを取得"""
    if items_per_page is None:
        items_per_page = get_items_per_page()
    
    start_idx = (page - 1) * items_per_page
    end_idx = start_idx + items_per_page
    return items[start_idx:end_idx]

def get_page_range(total_items: int, page: int, items_per_page: int = None) -> Tuple[int, int]:
    """現在ページの表示範囲を取得"""
    if items_per_page is None:
        items_per_page = get_items_per_page()
    
    start_item = (page - 1) * items_per_page + 1
    end_item = min(page * items_per_page, total_items)
    return start_item, end_item

# =============================================================================
# ファイル処理
# =============================================================================

def load_excel_file(uploaded_file, file_identifier: str) -> Tuple[Optional[pd.DataFrame], str, str]:
    """Excelファイル読み込み"""
    try:
        # ファイル検証
        is_valid, error_msg = validate_file(uploaded_file)
        if not is_valid:
            st.error(f"❌ {error_msg}")
            return None, "", ""
        
        # Excelファイル読み込み
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
        
        if not sheet_name:
            return None, uploaded_file.name, ""
        
        # データ読み込み
        data = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str).fillna("")
        
        # データサイズ制限
        if len(data) > CONFIG.MAX_DATA_ROWS:
            st.warning(f"⚠️ データが大きすぎます。最初の{CONFIG.MAX_DATA_ROWS:,}行のみ読み込まれました。")
            data = data.head(CONFIG.MAX_DATA_ROWS)
        
        st.success(f"✅ 「{sheet_name}」から {len(data):,} 行、{len(data.columns)} 列を読み込みました")
        
        # データプレビュー
        with st.expander("📊 データプレビュー"):
            st.dataframe(data.head(10), use_container_width=True)
        
        return data, uploaded_file.name, sheet_name
        
    except Exception as e:
        st.error(f"ファイル読み込みエラー: {str(e)}")
        logger.error(f"File loading error: {e}")
        return None, "", ""

# =============================================================================
# 在庫比較エンジン
# =============================================================================

class InventoryComparator:
    """在庫比較エンジン"""
    
    def __init__(self):
        self.key_columns = CONFIG.KEY_COLUMNS
        self.stock_column = CONFIG.STOCK_COLUMN
    
    def validate_data(self, data1: pd.DataFrame, data2: pd.DataFrame) -> Tuple[bool, str]:
        """データ検証"""
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
    
    def _generate_key(self, row: pd.Series) -> str:
        """行の一意キーを生成"""
        return "|".join([str(row.get(col, "")).strip() for col in self.key_columns])
    
    def _create_data_map(self, data: pd.DataFrame) -> Dict[str, Dict]:
        """データをマップに変換"""
        return {
            self._generate_key(row): row.to_dict() 
            for _, row in data.iterrows() 
            if self._generate_key(row)
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
    
    def compare(self, data1: pd.DataFrame, data2: pd.DataFrame) -> Tuple[List[ComparisonItem], ComparisonSummary]:
        """在庫比較実行"""
        # データをマップに変換
        map1 = self._create_data_map(data1)
        map2 = self._create_data_map(data2)
        
        all_keys = sorted(set(map1.keys()) | set(map2.keys()))
        
        items = []
        summary = ComparisonSummary()
        
        for key in all_keys:
            row1 = map1.get(key)
            row2 = map2.get(key)
            
            # 在庫値を取得
            stock1_numeric, stock1_display = (
                parse_stock_value(row1.get(self.stock_column)) if row1 else (0.0, "")
            )
            stock2_numeric, stock2_display = (
                parse_stock_value(row2.get(self.stock_column)) if row2 else (0.0, "")
            )
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
                
                items.append(ComparisonItem(
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
        
        return items, summary

# =============================================================================
# CSV出力
# =============================================================================

def create_csv_data(items: List[ComparisonItem], columns: List[str], encoding_choice: str) -> bytes:
    """CSV出力データ作成"""
    if not items:
        return b""
    
    export_data = []
    for item in items:
        stock1_display, stock2_display, stock_change_display = format_stock_display(item)
        
        row = {
            '変更タイプ': CONFIG.TYPE_EXPORT_MAP[item.type],
            'ファイル1在庫': stock1_display,
            'ファイル2在庫': stock2_display,
            '在庫変化': stock_change_display,
            '比較キー': item.key,
        }
        
        # 元の列データを追加
        for col in columns:
            row[col] = item.data.get(col, '')
        
        export_data.append(row)
    
    df = pd.DataFrame(export_data)
    encoding = CONFIG.ENCODING_MAP.get(encoding_choice, "utf-8-sig")
    return df.to_csv(index=False).encode(encoding)

# =============================================================================
# UI コンポーネント
# =============================================================================

def inject_responsive_css():
    """レスポンシブCSS注入"""
    st.markdown("""
    <style>
    /* ベースレスポンシブスタイル */
    .main > div {
        padding-top: 1rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    
    /* ヘッダーレスポンシブ */
    .responsive-header {
        background: linear-gradient(90deg, #4b6cb7 0%, #182848 100%);
        padding: 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        text-align: center;
    }
    
    .responsive-header h2 {
        color: white;
        margin: 0;
        font-size: clamp(1.2rem, 4vw, 2rem);
    }
    
    .responsive-header p {
        color: #E0E0E0;
        margin: 0.3rem 0 0 0;
        font-size: clamp(0.8rem, 2.5vw, 0.9rem);
    }
    
    /* ボタンレスポンシブ */
    .stButton > button {
        border-radius: 8px;
        font-weight: bold;
        width: 100%;
        padding: 0.5rem 1rem;
        font-size: clamp(0.8rem, 2.5vw, 1rem);
    }
    
    /* ファイル情報カード */
    .file-info-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
    }
    
    .file-info-card h4 {
        margin: 0 0 0.5rem 0;
        font-size: clamp(1rem, 3vw, 1.2rem);
    }
    
    .file-info-card p {
        margin: 0.2rem 0;
        font-size: clamp(0.8rem, 2.5vw, 0.9rem);
    }
    
    /* メトリクスカード */
    .metrics-container {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
        gap: 1rem;
        margin-bottom: 1rem;
    }
    
    .metric-card {
        background: white;
        border: 2px solid #e1e5e9;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        transition: transform 0.2s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    
    .metric-value {
        font-size: clamp(1.5rem, 4vw, 2rem);
        font-weight: bold;
        color: #2c3e50;
        margin: 0;
    }
    
    .metric-label {
        font-size: clamp(0.8rem, 2.5vw, 0.9rem);
        color: #7f8c8d;
        margin: 0.5rem 0 0 0;
    }
    
    /* ページネーション */
    .pagination-container {
        display: flex;
        flex-wrap: wrap;
        justify-content: space-between;
        align-items: center;
        gap: 0.5rem;
        margin: 1rem 0;
        padding: 1rem;
        background-color: #f8f9fa;
        border-radius: 8px;
    }
    
    .pagination-info {
        font-size: clamp(0.8rem, 2.5vw, 0.9rem);
        color: #6c757d;
    }
    
    /* モバイル最適化 */
    @media (max-width: 768px) {
        .main > div {
            padding-left: 0.5rem;
            padding-right: 0.5rem;
        }
        
        .responsive-header {
            padding: 0.8rem;
        }
        
        .file-info-card {
            padding: 0.8rem;
        }
        
        .metric-card {
            padding: 0.8rem;
        }
        
        .pagination-container {
            flex-direction: column;
            text-align: center;
        }
    }
    
    /* タブレット最適化 */
    @media (min-width: 769px) and (max-width: 1024px) {
        .metrics-container {
            grid-template-columns: repeat(2, 1fr);
        }
    }
    
    /* デスクトップ最適化 */
    @media (min-width: 1025px) {
        .metrics-container {
            grid-template-columns: repeat(4, 1fr);
        }
    }
    
    /* ダークモード対応 */
    @media (prefers-color-scheme: dark) {
        .file-info-card {
            background: linear-gradient(135deg, #4a5568 0%, #2d3748 100%);
        }
        
        .metric-card {
            background-color: #2d3748;
            border-color: #4a5568;
            color: #e2e8f0;
        }
        
        .metric-value {
            color: #e2e8f0;
        }
        
        .pagination-container {
            background-color: #2d3748;
        }
    }
    </style>
    """, unsafe_allow_html=True)

def render_header(comparison_completed: bool):
    """ヘッダー表示"""
    st.markdown("""
    <div class="responsive-header">
        <h2>📊 在庫差分比較ツール</h2>
        <p>Excelファイルの在庫データを比較し、差分を可視化するツールです。</p>
    </div>
    """, unsafe_allow_html=True)
    
    if comparison_completed and all(
        key in st.session_state for key in [
            SessionState.FILE1_NAME, SessionState.FILE1_SHEET,
            SessionState.FILE2_NAME, SessionState.FILE2_SHEET
        ]
    ):
        device_type = get_device_type()
        
        if device_type == 'mobile':
            # モバイルでは縦並び
            st.markdown(f"""
            <div class="file-info-card">
                <h4>📁 ファイル1（比較元）</h4>
                <p><strong>ファイル名:</strong> {st.session_state[SessionState.FILE1_NAME]}</p>
                <p><strong>シート:</strong> {st.session_state[SessionState.FILE1_SHEET]}</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div class="file-info-card">
                <h4>📁 ファイル2（比較先）</h4>
                <p><strong>ファイル名:</strong> {st.session_state[SessionState.FILE2_NAME]}</p>
                <p><strong>シート:</strong> {st.session_state[SessionState.FILE2_SHEET]}</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            # タブレット・デスクトップでは横並び
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"""
                <div class="file-info-card">
                    <h4>📁 ファイル1（比較元）</h4>
                    <p><strong>ファイル名:</strong> {st.session_state[SessionState.FILE1_NAME]}</p>
                    <p><strong>シート:</strong> {st.session_state[SessionState.FILE1_SHEET]}</p>
                </div>
                """, unsafe_allow_html=True)
            with col2:
                st.markdown(f"""
                <div class="file-info-card">
                    <h4>📁 ファイル2（比較先）</h4>
                    <p><strong>ファイル名:</strong> {st.session_state[SessionState.FILE2_NAME]}</p>
                    <p><strong>シート:</strong> {st.session_state[SessionState.FILE2_SHEET]}</p>
                </div>
                """, unsafe_allow_html=True)

def render_summary_metrics(summary: ComparisonSummary):
    """サマリーメトリクス表示"""
    st.markdown(f"""
    <div class="metrics-container">
        <div class="metric-card">
            <div class="metric-value">{summary.added:,}</div>
            <div class="metric-label">➕ 追加</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{summary.deleted:,}</div>
            <div class="metric-label">➖ 削除</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{summary.modified:,}</div>
            <div class="metric-label">🔄 在庫変更</div>
        </div>
        <div class="metric-card">
            <div class="metric-value">{summary.unchanged:,}</div>
            <div class="metric-label">✅ 変更なし</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

def render_sidebar(comparison_completed: bool):
    """サイドバー表示"""
    with st.sidebar:
        # デバイスタイプ選択（開発用）
        st.markdown("### 📱 表示モード")
        device_type = st.selectbox(
            "デバイスタイプ",
            ["desktop", "tablet", "mobile"],
            index=["desktop", "tablet", "mobile"].index(st.session_state.get('device_type', 'desktop')),
            help="表示を確認するデバイスタイプを選択"
        )
        if device_type != st.session_state.get('device_type'):
            st.session_state.device_type = device_type
            st.rerun()
        
        st.markdown("---")
        st.markdown("### ⚙️ 設定情報")
        
        with st.expander("比較設定", expanded=False):
            st.markdown("**比較項目:**")
            st.code(CONFIG.STOCK_COLUMN, language="text")
            st.markdown("**比較キー:**")
            st.code("\n".join(CONFIG.KEY_COLUMNS), language="text")
        
        st.markdown("---")
        st.markdown("### 📥 ダウンロード設定")
        encoding = st.selectbox("文字コードを選択", ["UTF-8 (BOM付き)", "Shift_JIS"])
        
        # ダウンロードボタン
        if comparison_completed and SessionState.ALL_ITEMS in st.session_state:
            csv_data = create_csv_data(
                st.session_state[SessionState.ALL_ITEMS], 
                st.session_state[SessionState.ORIGINAL_COLUMNS], 
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

def render_file_upload_section() -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame], str, str, str, str]:
    """ファイルアップロードセクション表示"""
    device_type = get_device_type()
    
    if device_type == 'mobile':
        # モバイルでは縦並び
        st.subheader("📁 ファイル1（比較元）")
        file1 = st.file_uploader("Excelファイル1", type=['xlsx', 'xls'], key="file1")
        data1, file1_name, file1_sheet = load_excel_file(file1, "file1") if file1 else (None, "", "")
        
        st.markdown("---")
        
        st.subheader("📁 ファイル2（比較先）")
        file2 = st.file_uploader("Excelファイル2", type=['xlsx', 'xls'], key="file2")
        data2, file2_name, file2_sheet = load_excel_file(file2, "file2") if file2 else (None, "", "")
    else:
        # タブレット・デスクトップでは横並び
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📁 ファイル1（比較元）")
            file1 = st.file_uploader("Excelファイル1", type=['xlsx', 'xls'], key="file1")
            data1, file1_name, file1_sheet = load_excel_file(file1, "file1") if file1 else (None, "", "")
        
        with col2:
            st.subheader("📁 ファイル2（比較先）")
            file2 = st.file_uploader("Excelファイル2", type=['xlsx', 'xls'], key="file2")
            data2, file2_name, file2_sheet = load_excel_file(file2, "file2") if file2 else (None, "", "")
    
    return data1, data2, file1_name, file1_sheet, file2_name, file2_sheet

def render_pagination_controls(total_items: int, current_page: int, tab_key: str) -> int:
    """ページング制御表示"""
    items_per_page = get_items_per_page()
    device_type = get_device_type()
    
    if total_items <= items_per_page:
        return current_page
    
    total_pages, max_page, _, _ = get_page_info(total_items, items_per_page)
    start_item, end_item = get_page_range(total_items, current_page, items_per_page)
    
    # ページング情報の表示
    st.markdown(f"""
    <div class="pagination-container">
        <div class="pagination-info">
            表示範囲: {start_item:,} - {end_item:,} 件 / 全 {total_items:,} 件 (ページ {current_page} / {total_pages})
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    if device_type == 'mobile':
        # モバイルでは縦並び
        new_page = st.selectbox(
            "ページ選択",
            range(1, total_pages + 1),
            index=current_page - 1,
            key=f"page_select_{tab_key}"
        )
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("◀️ 前のページ", key=f"prev_{tab_key}", disabled=(current_page == 1), use_container_width=True):
                new_page = max(1, current_page - 1)
        
        with col2:
            if st.button("▶️ 次のページ", key=f"next_{tab_key}", disabled=(current_page == total_pages), use_container_width=True):
                new_page = min(total_pages, current_page + 1)
    else:
        # タブレット・デスクトップでは横並び
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            if st.button("⏮️ 最初", key=f"first_{tab_key}", disabled=(current_page == 1)):
                new_page = 1
        
        with col2:
            if st.button("◀️ 前", key=f"prev_{tab_key}", disabled=(current_page == 1)):
                new_page = max(1, current_page - 1)
        
        with col3:
            new_page = st.selectbox(
                "ページ",
                range(1, total_pages + 1),
                index=current_page - 1,
                key=f"page_select_{tab_key}",
                label_visibility="collapsed"
            )
        
        with col4:
            if st.button("▶️ 次", key=f"next_{tab_key}", disabled=(current_page == total_pages)):
                new_page = min(total_pages, current_page + 1)
        
        with col5:
            if st.button("⏭️ 最後", key=f"last_{tab_key}", disabled=(current_page == total_pages)):
                new_page = total_pages
    
    return new_page

def render_results(items: List[ComparisonItem], columns: List[str], summary: ComparisonSummary):
    """結果表示"""
    if not items:
        st.info("🎉 比較対象のアイテムが見つかりませんでした")
        return
    
    st.markdown("## 📋 比較結果")
    
    # サマリーメトリクス表示
    render_summary_metrics(summary)
    
    # タブ設定
    tab_configs = [
        (f"🔍 全て ({summary.total_items:,})", "all"),
        (f"➕ 追加 ({summary.added:,})", "added"),
        (f"➖ 削除 ({summary.deleted:,})", "deleted"),
        (f"🔄 変更 ({summary.modified:,})", "modified"),
        (f"✅ 同じ ({summary.unchanged:,})", "unchanged")
    ]
    
    tabs = st.tabs([config[0] for config in tab_configs])
    
    for tab, (_, filter_type) in zip(tabs, tab_configs):
        with tab:
            if filter_type == "all":
                filtered_items = items
            else:
                filtered_items = [item for item in items if item.type == filter_type]
            
            if not filtered_items:
                st.info("該当するアイテムはありません")
                continue
            
            # ページング状態の管理
            page_key = f"page_{filter_type}"
            if page_key not in st.session_state:
                st.session_state[page_key] = 1
            
            # ページング制御
            current_page = render_pagination_controls(
                len(filtered_items), 
                st.session_state[page_key], 
                filter_type
            )
            
            # ページが変更された場合は更新
            if current_page != st.session_state[page_key]:
                st.session_state[page_key] = current_page
                st.rerun()
            
            # 現在ページのアイテムを取得
            page_items = get_page_items(filtered_items, current_page)
            
            if not page_items:
                st.info("このページには表示するアイテムがありません")
                continue
            
            # 表示データの準備
            display_data = []
            for item in page_items:
                stock1_display, stock2_display, stock_change_display = format_stock_display(item)
                
                row = {
                    '変更タイプ': CONFIG.TYPE_DISPLAY_MAP[item.type],
                    '在庫(元)': stock1_display,
                    '在庫(先)': stock2_display,
                    '在庫変化': stock_change_display
                }
                
                # 元の列データを追加
                for col in columns:
                    row[col] = item.data.get(col, '')
                
                display_data.append(row)
            
            df = pd.DataFrame(display_data)
            
            # デバイスタイプに応じた表示調整
            device_type = get_device_type()
            height = 400 if device_type == 'mobile' else min(600, len(df) * 35 + 38)
            
            st.dataframe(
                df, 
                use_container_width=True, 
                hide_index=True, 
                height=height
            )
            
            # 現在の表示情報
            items_per_page = get_items_per_page()
            start_item, end_item = get_page_range(len(filtered_items), current_page, items_per_page)
            st.caption(f"現在のページ: {start_item:,} - {end_item:,} 件 / 全 {len(filtered_items):,} 件")

# =============================================================================
# メインアプリケーション
# =============================================================================

def setup_page():
    """ページ設定"""
    st.set_page_config(
        page_title="在庫差分比較ツール v7.1 (Refactored)",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    inject_responsive_css()

def handle_comparison_execution(comparator: InventoryComparator, data1: pd.DataFrame, data2: pd.DataFrame, 
                               file1_name: str, file1_sheet: str, file2_name: str, file2_sheet: str):
    """比較実行処理"""
    is_valid, error_msg = comparator.validate_data(data1, data2)
    
    if not is_valid:
        st.error(f"❌ データ検証エラー: {error_msg}")
        return
    
    if st.button("🔍 比較実行", type="primary", use_container_width=True):
        with st.spinner("比較処理を実行中..."):
            try:
                items, summary = comparator.compare(data1, data2)
                
                # 結果を保存
                SessionState.save_comparison_result(
                    items, summary, data1.columns.tolist(),
                    file1_name, file1_sheet, file2_name, file2_sheet
                )
                
                st.success("✅ 比較処理が完了しました！")
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ 比較処理中にエラーが発生しました: {str(e)}")
                logger.error(f"Comparison error: {e}")

def main():
    """メイン関数"""
    # 初期設定
    setup_page()
    SessionState.initialize()
    
    # 比較エンジン初期化
    comparator = InventoryComparator()
    
    # UI描画
    comparison_completed = st.session_state[SessionState.COMPARISON_COMPLETED]
    render_header(comparison_completed)
    render_sidebar(comparison_completed)
    
    # メインコンテンツ
    if not comparison_completed:
        # ファイルアップロードと比較実行
        data1, data2, file1_name, file1_sheet, file2_name, file2_sheet = render_file_upload_section()
        
        if data1 is not None and data2 is not None:
            handle_comparison_execution(comparator, data1, data2, file1_name, file1_sheet, file2_name, file2_sheet)
    else:
        # クリアボタン
        if st.button("🗑️ クリア", type="secondary", use_container_width=True):
            SessionState.clear_comparison_data()
            st.rerun()
        
        # 結果表示
        render_results(
            st.session_state[SessionState.ALL_ITEMS], 
            st.session_state[SessionState.ORIGINAL_COLUMNS], 
            st.session_state[SessionState.SUMMARY]
        )

if __name__ == "__main__":
    main()