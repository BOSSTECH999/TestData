"""
数智质控 v0.5
检测数据智能分析与质控预警系统
功能：
1. 用户上传Excel，智能识别数据方向（列=项目 or 行=项目）
2. 用户选择分类字段进行分组建模
3. 每个项目可设置小数位
4. 模型管理（多模型保存/加载/删除）
5. 数据智能分析 + 异常报告
6. 相关性分析：热力图、散点图、强相关项目对列表
7. 数据量控制（滑动窗口）
8. 模型滚动更新
"""

import streamlit as st
import pandas as pd
import numpy as np
from scipy import stats
import plotly.express as px
import plotly.graph_objects as go
import json
import os
import hashlib
import time
import base64
import io
from datetime import datetime

# ========== 页面配置 ==========
st.set_page_config(
    page_title="数智质控 - 智能检测数据分析平台",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="auto"
)

# ========== 全局样式 ==========
st.markdown("""
<style>
    /* ===== 主题色 ===== */
    :root {
        --primary: #1a73e8;
        --primary-dark: #1557b0;
        --primary-light: #e8f0fe;
        --accent: #00c853;
        --warning: #ff9800;
        --danger: #d32f2f;
        --bg-dark: #0f1923;
        --bg-card: #1a2733;
        --text-primary: #ffffff;
        --text-secondary: #b0bec5;
        --border-color: #2a3a4a;
    }
    
    /* ===== 全局 ===== */
    .stApp {
        background: linear-gradient(135deg, #0f1923 0%, #1a2733 50%, #0f1923 100%);
    }
    
    /* ===== 侧边栏 ===== */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0d1520 0%, #162230 100%) !important;
        border-right: 1px solid var(--border-color);
    }
    section[data-testid="stSidebar"] .stRadio > label {
        color: var(--text-secondary) !important;
        font-size: 14px !important;
        padding: 10px 16px !important;
        border-radius: 8px !important;
        transition: all 0.3s ease !important;
    }
    section[data-testid="stSidebar"] .stRadio > label:hover {
        background: rgba(26, 115, 232, 0.1) !important;
        color: var(--text-primary) !important;
    }
    section[data-testid="stSidebar"] .stRadio > label[data-baseweb="radio--checked"],
    section[data-testid="stSidebar"] .stRadio [aria-checked="true"] ~ label {
        background: linear-gradient(90deg, rgba(26, 115, 232, 0.2), transparent) !important;
        color: var(--primary) !important;
        font-weight: 600 !important;
        border-left: 3px solid var(--primary) !important;
    }
    
    /* ===== 顶栏标题 ===== */
    .main-title {
        display: flex;
        align-items: center;
        gap: 16px;
        padding: 20px 0 10px 0;
        border-bottom: 1px solid var(--border-color);
        margin-bottom: 24px;
    }
    .main-title h1 {
        font-size: 28px !important;
        font-weight: 700 !important;
        background: linear-gradient(135deg, #4fc3f7, #1a73e8);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 0 !important;
    }
    .main-title .subtitle {
        font-size: 13px;
        color: var(--text-secondary);
        letter-spacing: 2px;
    }
    .main-title .logo-icon {
        font-size: 36px;
    }
    
    /* ===== 卡片 ===== */
    .card {
        background: var(--bg-card);
        border: 1px solid var(--border-color);
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 16px;
    }
    .card-title {
        font-size: 16px;
        font-weight: 600;
        color: var(--primary);
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    /* ===== 按钮 ===== */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--primary), var(--primary-dark)) !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 10px 24px !important;
        box-shadow: 0 4px 12px rgba(26, 115, 232, 0.3) !important;
        transition: all 0.3s ease !important;
    }
    .stButton > button[kind="primary"]:hover {
        box-shadow: 0 6px 20px rgba(26, 115, 232, 0.5) !important;
        transform: translateY(-1px);
    }
    
    /* ===== 指标卡片 ===== */
    [data-testid="stMetricValue"] {
        font-size: 28px !important;
        font-weight: 700 !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 13px !important;
        color: var(--text-secondary) !important;
    }
    [data-testid="stMetricDelta"] {
        font-size: 12px !important;
    }
    
    /* ===== 数据表格 ===== */
    .stDataFrame {
        border: 1px solid var(--border-color);
        border-radius: 8px;
        overflow: hidden;
    }
    
    /* ===== 上传区域 ===== */
    .stFileUploader {
        border: 2px dashed var(--border-color);
        border-radius: 12px;
        padding: 20px;
        background: rgba(26, 115, 232, 0.05);
    }
    
    /* ===== 成功/警告/错误 ===== */
    .stSuccess {
        background: rgba(0, 200, 83, 0.1) !important;
        border-left: 3px solid var(--accent) !important;
        border-radius: 8px !important;
    }
    .stWarning {
        background: rgba(255, 152, 0, 0.1) !important;
        border-left: 3px solid var(--warning) !important;
        border-radius: 8px !important;
    }
    .stError {
        background: rgba(211, 47, 47, 0.1) !important;
        border-left: 3px solid var(--danger) !important;
        border-radius: 8px !important;
    }
    .stInfo {
        background: rgba(26, 115, 232, 0.1) !important;
        border-left: 3px solid var(--primary) !important;
        border-radius: 8px !important;
    }
    
    /* ===== 移动端优化 ===== */
    @media (max-width: 768px) {
        .stButton > button {
            min-height: 44px;
            font-size: 16px;
        }
        .stNumberInput input {
            min-height: 44px;
            font-size: 16px;
        }
        .stSelectbox div[data-baseweb="select"] {
            min-height: 44px;
        }
        .main-title h1 {
            font-size: 22px !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# ========== 数据目录 ==========
DATA_DIR = "data"
MODEL_BASE_DIR = "models"
os.makedirs(DATA_DIR, exist_ok=True)


def get_model_dir(username):
    """获取用户专属模型目录"""
    model_dir = os.path.join(MODEL_BASE_DIR, username)
    os.makedirs(model_dir, exist_ok=True)
    return model_dir


# ========== 工具函数 ==========
def clean_columns(df):
    """清理列名"""
    df.columns = [str(col).strip().replace('\n', '').replace('\r', '') for col in df.columns]
    return df


def detect_numeric_cols(df, skip_keywords=None):
    """智能识别数值列"""
    if skip_keywords is None:
        skip_keywords = ["序号", "编号", "标识", "类型", "煤种", "时间", "日期", "备注", "样品", "名称", "品牌"]
    
    numeric_cols = []
    for col in df.columns:
        if any(kw in str(col) for kw in skip_keywords):
            continue
        test_series = pd.to_numeric(df[col], errors='coerce')
        valid_count = test_series.notna().sum()
        total_count = len(df[col].dropna())
        if total_count > 0 and valid_count / total_count > 0.4:
            numeric_cols.append(col)
    return numeric_cols


def get_classification_candidates(df, numeric_cols):
    """获取所有候选分类列，返回 {列名: 是否推荐}。让用户自选，系统仅提供推荐参考。"""
    # 关键词只用于推荐标记，不用于过滤
    recommend_keywords = ["煤种", "类型", "类别", "品种", "等级", "分类", "来源", "产地",
                          "供应商", "批次", "产线", "工段", "车间", "矿区", "品牌", "规格", "型号"]
    skip_keywords = ["序号", "编号", "标识", "时间", "日期", "备注", "样品", "名称", "电话", "地址", "联系人"]

    candidates = {}
    for col in df.columns:
        if col in numeric_cols:
            continue
        if any(kw in str(col) for kw in skip_keywords):
            continue

        # 推荐逻辑：关键词匹配 > 唯一值比例判断
        recommended = False
        if any(kw in str(col) for kw in recommend_keywords):
            recommended = True
        else:
            unique_vals = df[col].dropna().nunique()
            total_vals = len(df[col].dropna())
            if total_vals > 0 and unique_vals / total_vals < 0.3 and unique_vals >= 2:
                recommended = True

        candidates[col] = recommended

    return candidates


def load_data(uploaded_file):
    """读取Excel，返回原始数据和相关信息"""
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    df = clean_columns(df)
    return df


def transpose_if_needed(df):
    """检测数据方向：如果第一列都是文本且列名像数据项，则转置"""
    # 检查是否需要转置：如果行数远少于列数，且第一列值像检测项目名
    first_col = df.iloc[:, 0]
    numeric_ratio = pd.to_numeric(first_col, errors='coerce').notna().sum() / len(first_col.dropna()) if len(first_col.dropna()) > 0 else 1
    
    # 如果第一列大部分不是数字，且列数大于行数，可能需要转置
    if numeric_ratio < 0.3 and df.shape[1] > df.shape[0]:
        df_t = df.set_index(df.columns[0]).T.reset_index()
        df_t = df_t.rename(columns={df_t.columns[0]: "序号"})
        df_t = clean_columns(df_t)
        return df_t, True
    return df, False


def build_model(df_numeric, numeric_cols, group_col=None, group_value=None, decimal_settings=None, max_data_count=None):
    """从历史数据学习统计规律，支持滑动窗口控制数据量"""
    model = {}
    
    # 滑动窗口：只保留最近 max_data_count 条数据
    if max_data_count and len(df_numeric) > max_data_count:
        model["original_count"] = len(df_numeric)
        df_numeric = df_numeric.iloc[-max_data_count:]
        model["data_trimmed"] = True
    else:
        model["data_trimmed"] = False
    
    # 分组过滤
    if group_col and group_value:
        model["group_col"] = group_col
        model["group_value"] = str(group_value)
    
    # 1. 每个项目的统计特征
    model["item_stats"] = {}
    for col in numeric_cols:
        series = df_numeric[col].dropna()
        if len(series) < 3:
            continue
        
        # 获取小数位设置
        decimals = 3
        if decimal_settings and col in decimal_settings:
            decimals = decimal_settings[col]
        
        model["item_stats"][col] = {
            "mean": round(float(series.mean()), decimals),
            "std": round(float(series.std()), decimals),
            "median": round(float(series.median()), decimals),
            "q1": round(float(series.quantile(0.25)), decimals),
            "q3": round(float(series.quantile(0.75)), decimals),
            "iqr": round(float(series.quantile(0.75) - series.quantile(0.25)), decimals),
            "lower_fence": round(float(series.quantile(0.25) - 1.5 * (series.quantile(0.75) - series.quantile(0.25))), decimals),
            "upper_fence": round(float(series.quantile(0.75) + 1.5 * (series.quantile(0.75) - series.quantile(0.25))), decimals),
            "count": len(series),
            "decimals": decimals
        }
    
    # 2. 相关性
    valid_cols = [c for c in numeric_cols if c in model["item_stats"]]
    if len(valid_cols) >= 2:
        corr_matrix = df_numeric[valid_cols].corr()
        model["correlations"] = corr_matrix.to_dict()
        
        model["strong_pairs"] = []
        for i in range(len(valid_cols)):
            for j in range(i + 1, len(valid_cols)):
                r = corr_matrix.iloc[i, j]
                if abs(r) > 0.5 and not np.isnan(r):
                    model["strong_pairs"].append({
                        "col_a": valid_cols[i],
                        "col_b": valid_cols[j],
                        "r": round(float(r), 3),
                        "direction": "正相关" if r > 0 else "负相关"
                    })
        model["strong_pairs"].sort(key=lambda x: abs(x["r"]), reverse=True)
    
    # 3. 回归模型
    model["regression_models"] = {}
    for pair in model.get("strong_pairs", []):
        col_a, col_b = pair["col_a"], pair["col_b"]
        mask = df_numeric[col_a].notna() & df_numeric[col_b].notna()
        if mask.sum() < 5:
            continue
        x = df_numeric.loc[mask, col_a].values
        y = df_numeric.loc[mask, col_b].values
        slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
        residuals = y - (slope * x + intercept)
        residual_std = float(residuals.std())
        model["regression_models"][f"{col_a}|{col_b}"] = {
            "slope": float(slope),
            "intercept": float(intercept),
            "r_squared": round(float(r_value ** 2), 4),
            "residual_std": round(residual_std, 4)
        }
        if abs(slope) > 1e-10:
            model["regression_models"][f"{col_b}|{col_a}"] = {
                "slope": float(1/slope),
                "intercept": float(-intercept/slope),
                "r_squared": round(float(r_value ** 2), 4),
                "residual_std": round(residual_std / abs(float(slope)), 4)
            }
    
    model["trained_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    model["sample_count"] = len(df_numeric)
    model["numeric_cols"] = numeric_cols
    
    return model


def check_anomaly(new_row, model, history_df=None):
    """异常检测 + 极值标记"""
    anomalies = []
    numeric_cols = model.get("numeric_cols", list(model.get("item_stats", {}).keys()))
    
    # 1. 单项异常 + 极值标记
    for col in numeric_cols:
        if col not in model.get("item_stats", {}):
            continue
        val = new_row.get(col)
        if val is None or (isinstance(val, float) and np.isnan(val)):
            continue
        try:
            val = float(val)
        except:
            continue
            
        stats_info = model["item_stats"][col]
        z_score = abs(val - stats_info["mean"]) / stats_info["std"] if stats_info["std"] > 0 else 0
        decimals = stats_info.get("decimals", 3)
        
        if z_score > 3:
            direction = "严重偏高" if val > stats_info["mean"] else "严重偏低"
            anomalies.append({
                "项目": col,
                "类型": f"🔴 {direction}",
                "检测值": round(val, decimals),
                "正常范围": f"{round(stats_info['mean']-3*stats_info['std'], decimals)} ~ {round(stats_info['mean']+3*stats_info['std'], decimals)}",
                "偏离程度": f"Z={z_score:.1f}σ",
                "说明": f"偏离均值{z_score:.1f}个标准差"
            })
        elif val < stats_info["lower_fence"] or val > stats_info["upper_fence"]:
            direction = "偏高" if val > stats_info["upper_fence"] else "偏低"
            anomalies.append({
                "项目": col,
                "类型": f"🟡 {direction}",
                "检测值": round(val, decimals),
                "正常范围": f"{round(stats_info['lower_fence'], decimals)} ~ {round(stats_info['upper_fence'], decimals)}",
                "偏离程度": f"Z={z_score:.1f}σ",
                "说明": "超出IQR围栏"
            })
        
        # 极值标记：当前值是历史最大/最小值
        if history_df is not None and col in history_df.columns:
            series = pd.to_numeric(history_df[col], errors='coerce').dropna()
            if len(series) > 0:
                hist_max = float(series.max())
                hist_min = float(series.min())
                if val == hist_max and val > stats_info["mean"]:
                    anomalies.append({
                        "项目": col,
                        "类型": "⬆️ 创历史新高",
                        "检测值": round(val, decimals),
                        "正常范围": f"历史范围：{round(hist_min, decimals)} ~ {round(hist_max, decimals)}",
                        "偏离程度": f"均值={stats_info['mean']}",
                        "说明": "当前值为历史数据中的最大值，可能代表新工况，需关注"
                    })
                elif val == hist_min and val < stats_info["mean"]:
                    anomalies.append({
                        "项目": col,
                        "类型": "⬇️ 创历史新低",
                        "检测值": round(val, decimals),
                        "正常范围": f"历史范围：{round(hist_min, decimals)} ~ {round(hist_max, decimals)}",
                        "偏离程度": f"均值={stats_info['mean']}",
                        "说明": "当前值为历史数据中的最小值，可能代表新工况，需关注"
                    })
    
    # 2. 逻辑异常
    for pair in model.get("strong_pairs", []):
        col_a, col_b = pair["col_a"], pair["col_b"]
        val_a, val_b = new_row.get(col_a), new_row.get(col_b)
        if val_a is None or val_b is None:
            continue
        try:
            val_a, val_b = float(val_a), float(val_b)
        except:
            continue
        if np.isnan(val_a) or np.isnan(val_b):
            continue
        
        key = f"{col_a}|{col_b}"
        if key not in model.get("regression_models", {}):
            continue
        reg = model["regression_models"][key]
        predicted_b = reg["slope"] * val_a + reg["intercept"]
        residual = abs(val_b - predicted_b)
        
        if reg["residual_std"] > 0 and residual > 2.5 * reg["residual_std"]:
            dec = model["item_stats"].get(col_b, {}).get("decimals", 3)
            updown = "应该同步变化" if pair["direction"] == "正相关" else "应该此消彼长"
            rule = f"按历史规律，{col_a}涨{col_b}也涨" if pair["direction"] == "正相关" else f"按历史规律，{col_a}越高{col_b}越低"
            anomalies.append({
                "项目": col_a,
                "类型": f"🔵 {col_a}与{col_b}搭配异常",
                "检测值": f"{col_a}={round(val_a, dec)}, {col_b}={round(val_b, dec)}",
                "正常范围": f"{col_b}应在{round(predicted_b-2.5*reg['residual_std'], dec)} ~ {round(predicted_b+2.5*reg['residual_std'], dec)}",
                "偏离程度": f"{residual:.1f}",
                "说明": f"{rule}，但当前{col_a}={round(val_a, dec)}时{col_b}={round(val_b, dec)}，{updown}才对"
            })
    
    return anomalies


def save_model(model, name, username):
    """保存模型到用户专属目录"""
    safe_name = name.replace("/", "_").replace("\\", "_").replace(" ", "_")
    model_dir = get_model_dir(username)
    filepath = os.path.join(model_dir, f"{safe_name}.json")
    # 在模型内保存友好名称
    model["display_name"] = name
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(model, f, ensure_ascii=False, indent=2)
    return filepath


def load_model_file(filepath):
    """从文件加载模型"""
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)


def get_saved_models(username=None):
    """获取当前用户已保存的模型"""
    models = []
    if not username:
        return models
    model_dir = get_model_dir(username)
    if os.path.exists(model_dir):
        for f in os.listdir(model_dir):
            if f.endswith(".json"):
                filepath = os.path.join(model_dir, f)
                try:
                    m = load_model_file(filepath)
                    # 优先使用模型内的友好名称
                    display_name = m.get("display_name", f.replace(".json", ""))
                    group_val = m.get("group_value", "全部")
                    models.append({
                        "filename": f,
                        "filepath": filepath,
                        "name": f.replace(".json", ""),
                        "display_name": display_name,
                        "group": group_val,
                        "sample_count": m.get("sample_count", 0),
                        "trained_at": m.get("trained_at", ""),
                        "item_count": len(m.get("item_stats", {})),
                        "pair_count": len(m.get("strong_pairs", []))
                    })
                except:
                    pass
    return models


def get_model_display_name(model_name, model=None, username=None):
    """获取模型的友好显示名称"""
    if model and model.get("display_name"):
        return model["display_name"]
    # 如果有分组信息，用分组信息生成友好名称
    if model and model.get("group_value") and model.get("group_value") != "全部":
        group_col = model.get("group_col", "分类")
        return f"{group_col}：{model['group_value']}"
    # 从已保存模型中查找
    saved = get_saved_models(username)
    for sm in saved:
        if sm["name"] == model_name:
            if sm.get("display_name") and sm["display_name"] != model_name:
                return sm["display_name"]
            if sm["group"] != "全部":
                return sm["group"]
    # 兜底：去掉 model_ 前缀和 _latest 后缀
    clean = model_name.replace("model_", "").replace("_latest", "").replace("_all", "全部数据")
    if clean:
        return clean
    return "全部数据"


# ========== 初始化session state ==========
if "df_raw" not in st.session_state:
    st.session_state.df_raw = None
if "models" not in st.session_state:
    st.session_state.models = {}
if "max_data_count" not in st.session_state:
    st.session_state.max_data_count = 500
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "current_user" not in st.session_state:
    st.session_state.current_user = None
if "login_time" not in st.session_state:
    st.session_state.login_time = None
if "pending_2fa_user" not in st.session_state:
    st.session_state.pending_2fa_user = None
if "new_data" not in st.session_state:
    st.session_state.new_data = None  # 暂存手动输入记录

# ========== 用户管理 ==========
USER_DIR = "users"
os.makedirs(USER_DIR, exist_ok=True)
SESSION_TIMEOUT = 7200  # 2小时，单位秒
LOCKOUT_ATTEMPTS = 5     # 5次失败后锁定
LOCKOUT_DURATION = 900   # 锁定15分钟，单位秒

def generate_salt():
    """生成随机盐值"""
    return os.urandom(16).hex()

def hash_password(password, salt):
    """密码加盐哈希"""
    return hashlib.sha256((salt + password).encode('utf-8')).hexdigest()

def get_user_file(username):
    """获取用户数据文件路径"""
    safe_name = hashlib.md5(username.encode('utf-8')).hexdigest()[:16]
    return os.path.join(USER_DIR, f"{safe_name}.json")

def is_first_user():
    """检查是否还没有任何用户注册"""
    if not os.path.exists(USER_DIR):
        return True
    return len([f for f in os.listdir(USER_DIR) if f.endswith(".json")]) == 0

def get_all_users():
    """获取所有用户信息（管理员用）"""
    users = []
    if not os.path.exists(USER_DIR):
        return users
    for f in os.listdir(USER_DIR):
        if f.endswith(".json"):
            filepath = os.path.join(USER_DIR, f)
            try:
                with open(filepath, "r", encoding="utf-8") as fp:
                    user_data = json.load(fp)
                users.append(user_data)
            except:
                pass
    return users

def register_user(username, password, email):
    """注册新用户"""
    if not username or not password:
        return False, "用户名和密码不能为空"
    if len(username) < 3:
        return False, "用户名至少3个字符"
    if len(password) < 6:
        return False, "密码至少6个字符"
    if not email or "@" not in email:
        return False, "请输入有效的邮箱地址"
    
    user_file = get_user_file(username)
    if os.path.exists(user_file):
        return False, "用户名已被注册"
    
    salt = generate_salt()
    # 第一个注册的用户自动成为管理员
    role = "admin" if is_first_user() else "user"
    
    user_data = {
        "username": username,
        "password_hash": hash_password(password, salt),
        "salt": salt,
        "email": email,
        "role": role,
        "totp_enabled": False,
        "totp_secret": "",
        "failed_attempts": 0,
        "locked_until": 0,
        "banned": False,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "last_login": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    with open(user_file, "w", encoding="utf-8") as f:
        json.dump(user_data, f, ensure_ascii=False, indent=2)
    
    role_hint = "（管理员）" if role == "admin" else ""
    return True, f"注册成功{role_hint}"

def load_user(username):
    """加载用户数据"""
    user_file = get_user_file(username)
    if not os.path.exists(user_file):
        return None
    with open(user_file, "r", encoding="utf-8") as f:
        return json.load(f)

def save_user(user_data):
    """保存用户数据"""
    user_file = get_user_file(user_data["username"])
    with open(user_file, "w", encoding="utf-8") as f:
        json.dump(user_data, f, ensure_ascii=False, indent=2)

def verify_user(username, password):
    """验证用户登录（含锁定检查）"""
    user_data = load_user(username)
    if user_data is None:
        return False, "用户名不存在"
    
    # 检查是否被封禁
    if user_data.get("banned", False):
        return False, "账号已被封禁，请联系管理员"
    
    # 检查是否被锁定
    now = time.time()
    if user_data.get("locked_until", 0) > now:
        remaining = int(user_data["locked_until"] - now)
        mins = remaining // 60 + 1
        return False, f"账号已锁定，请{mins}分钟后再试"
    
    # 如果锁定时间已过，解锁
    if user_data.get("locked_until", 0) > 0 and user_data["locked_until"] <= now:
        user_data["failed_attempts"] = 0
        user_data["locked_until"] = 0
    
    # 验证密码（兼容旧版无盐密码）
    salt = user_data.get("salt", "")
    if salt:
        pw_hash = hash_password(password, salt)
    else:
        # 旧版无盐密码，向后兼容
        pw_hash = hashlib.sha256(password.encode('utf-8')).hexdigest()
    
    if user_data["password_hash"] != pw_hash:
        # 密码错误，增加失败次数
        user_data["failed_attempts"] = user_data.get("failed_attempts", 0) + 1
        if user_data["failed_attempts"] >= LOCKOUT_ATTEMPTS:
            user_data["locked_until"] = now + LOCKOUT_DURATION
            save_user(user_data)
            return False, f"连续{LOCKOUT_ATTEMPTS}次密码错误，账号锁定{LOCKOUT_DURATION // 60}分钟"
        save_user(user_data)
        remaining = LOCKOUT_ATTEMPTS - user_data["failed_attempts"]
        return False, f"密码错误，还剩{remaining}次尝试机会"
    
    # 密码正确，重置失败次数
    user_data["failed_attempts"] = 0
    user_data["locked_until"] = 0
    
    # 旧版无盐密码自动升级为加盐密码
    if not salt:
        new_salt = generate_salt()
        user_data["salt"] = new_salt
        user_data["password_hash"] = hash_password(password, new_salt)
    
    # 更新最后登录时间
    user_data["last_login"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_user(user_data)
    return True, user_data

def reset_password(username, email, new_password):
    """通过邮箱验证重置密码"""
    user_data = load_user(username)
    if user_data is None:
        return False, "用户名不存在"
    if user_data.get("email", "").lower() != email.lower():
        return False, "邮箱与注册时不匹配"
    if len(new_password) < 6:
        return False, "新密码至少6个字符"
    salt = user_data.get("salt", generate_salt())
    user_data["salt"] = salt
    user_data["password_hash"] = hash_password(new_password, salt)
    user_data["failed_attempts"] = 0
    user_data["locked_until"] = 0
    save_user(user_data)
    return True, "密码重置成功"

def verify_totp(totp_secret, code):
    """验证TOTP动态口令"""
    try:
        import pyotp
        totp = pyotp.TOTP(totp_secret)
        return totp.verify(code, valid_window=1)  # 允许前后30秒的窗口
    except ImportError:
        # pyotp未安装时，用标准库实现
        return _totp_verify_fallback(totp_secret, code)

def _totp_verify_fallback(secret, code):
    """TOTP备用验证（不依赖pyotp）"""
    try:
        import hmac
        import struct
        key = base64.b32decode(secret, casefold=True)
        now = int(time.time()) // 30
        for offset in [-1, 0, 1]:  # 允许前后30秒
            t = now + offset
            msg = struct.pack(">Q", t)
            h = hmac.new(key, msg, hashlib.sha1).digest()
            o = h[19] & 15
            token = (struct.unpack(">I", h[o:o+4])[0] & 0x7fffffff) % 1000000
            if str(token).zfill(6) == code:
                return True
        return False
    except:
        return False

def generate_totp_secret():
    """生成TOTP密钥"""
    try:
        import pyotp
        return pyotp.random_base32()
    except ImportError:
        # 备用：生成随机base32密钥
        return base64.b32encode(os.urandom(20)).decode('utf-8')

def get_totp_uri(username, secret):
    """生成TOTP URI（用于扫码绑定）"""
    return f"otpauth://totp/数智质控:{username}?secret={secret}&issuer=数智质控"

def get_totp_qr_base64(username, secret):
    """生成TOTP二维码的base64图片"""
    try:
        import qrcode
        uri = get_totp_uri(username, secret)
        qr = qrcode.QRCode(version=1, box_size=8, border=2)
        qr.add_data(uri)
        qr.make(fit=True)
        img = qr.make_image(fill_color="white", back_color="black")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode('utf-8')
    except ImportError:
        return None


# ========== 会话超时检查 ==========
if st.session_state.logged_in and st.session_state.login_time:
    elapsed = time.time() - st.session_state.login_time
    if elapsed > SESSION_TIMEOUT:
        st.session_state.logged_in = False
        st.session_state.current_user = None
        st.session_state.login_time = None
        st.session_state.pending_2fa_user = None
        st.warning("⏰ 登录已超时，请重新登录")


# ========== 登录/注册界面 ==========
if not st.session_state.logged_in:
    # 封面样式
    st.markdown("""
    <style>
        .login-container {
            max-width: 420px;
            margin: 0 auto;
            padding: 20px;
        }
        .brand-area {
            text-align: center;
            padding: 40px 0 30px 0;
        }
        .brand-logo {
            font-size: 64px;
            margin-bottom: 10px;
        }
        .brand-name {
            font-size: 40px;
            font-weight: 800;
            background: linear-gradient(135deg, #4fc3f7, #1a73e8, #7c4dff);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            letter-spacing: 4px;
            margin-bottom: 8px;
        }
        .brand-subtitle {
            font-size: 14px;
            color: #607d8b;
            letter-spacing: 6px;
            text-transform: uppercase;
        }
        .brand-tagline {
            margin-top: 16px;
            font-size: 15px;
            color: #78909c;
            line-height: 1.8;
        }
        .feature-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
            margin-top: 24px;
        }
        .feature-item {
            background: rgba(26, 115, 232, 0.08);
            border: 1px solid rgba(26, 115, 232, 0.15);
            border-radius: 10px;
            padding: 16px 14px;
            text-align: center;
        }
        .feature-icon {
            font-size: 24px;
            margin-bottom: 6px;
        }
        .feature-name {
            font-size: 13px;
            color: #b0bec5;
            font-weight: 500;
        }
        .login-card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 16px;
            padding: 32px 28px;
            margin-top: 20px;
        }
        .login-title {
            font-size: 20px;
            font-weight: 700;
            color: #ffffff;
            text-align: center;
            margin-bottom: 24px;
        }
        .powered-by {
            text-align: center;
            margin-top: 32px;
            padding: 20px 0;
            color: #455a64;
            font-size: 12px;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # 品牌展示区
    st.markdown("""
    <div class="brand-area">
        <div class="brand-logo">📊</div>
        <div class="brand-name">数智质控</div>
        <div class="brand-subtitle">Smart Quality Control Platform</div>
        <div class="brand-tagline">
            检测数据智能分析 · 异常自动识别 · 质控实时预警
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 功能亮点
    st.markdown("""
    <div class="feature-grid">
        <div class="feature-item">
            <div class="feature-icon">🧠</div>
            <div class="feature-name">智能建模分析</div>
        </div>
        <div class="feature-item">
            <div class="feature-icon">🔍</div>
            <div class="feature-name">异常自动检测</div>
        </div>
        <div class="feature-item">
            <div class="feature-icon">📈</div>
            <div class="feature-name">趋势诊断预警</div>
        </div>
        <div class="feature-item">
            <div class="feature-icon">🔗</div>
            <div class="feature-name">项目相关性发现</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 登录/注册卡片
    st.markdown("""<div class="login-card">""", unsafe_allow_html=True)
    
    auth_tab = st.tabs(["🔐 登录", "📝 注册", "🔑 找回密码"])
    
    with auth_tab[0]:
        st.markdown('<div class="login-title">欢迎回来</div>', unsafe_allow_html=True)
        
        # 2FA验证步骤
        if st.session_state.get("pending_2fa_user"):
            st.info(f"🔐 用户 {st.session_state.pending_2fa_user} 已开启动态口令验证")
            totp_code = st.text_input("动态口令（6位数字）", placeholder="打开验证器App查看", key="totp_code")
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("✅ 验证", type="primary", use_container_width=True):
                    if not totp_code:
                        st.warning("请输入动态口令")
                    else:
                        user_data = load_user(st.session_state.pending_2fa_user)
                        if user_data and verify_totp(user_data.get("totp_secret", ""), totp_code):
                            st.session_state.logged_in = True
                            st.session_state.current_user = st.session_state.pending_2fa_user
                            st.session_state.login_time = time.time()
                            st.session_state.pending_2fa_user = None
                            st.rerun()
                        else:
                            st.error("动态口令错误")
            with col_b:
                if st.button("❌ 取消", use_container_width=True):
                    st.session_state.pending_2fa_user = None
                    st.rerun()
        else:
            login_user = st.text_input("用户名", placeholder="请输入用户名", key="login_user")
            login_pass = st.text_input("密码", type="password", placeholder="请输入密码", key="login_pass")
            if st.button("🚀 登录", type="primary", use_container_width=True):
                if not login_user or not login_pass:
                    st.warning("请输入用户名和密码")
                else:
                    ok, result = verify_user(login_user, login_pass)
                    if ok:
                        # 检查是否开启2FA
                        if result.get("totp_enabled", False):
                            st.session_state.pending_2fa_user = login_user
                            st.rerun()
                        else:
                            st.session_state.logged_in = True
                            st.session_state.current_user = login_user
                            st.session_state.login_time = time.time()
                            st.rerun()
                    else:
                        st.error(result)
    
    with auth_tab[1]:
        st.markdown('<div class="login-title">创建账号</div>', unsafe_allow_html=True)
        reg_user = st.text_input("用户名", placeholder="3个字符以上", key="reg_user")
        reg_pass = st.text_input("密码", type="password", placeholder="6个字符以上", key="reg_pass")
        reg_pass2 = st.text_input("确认密码", type="password", placeholder="再次输入密码", key="reg_pass2")
        reg_email = st.text_input("找回邮箱", placeholder="用于密码找回，不收集其他信息", key="reg_email")
        if st.button("✅ 注册", type="primary", use_container_width=True):
            if not reg_user or not reg_pass or not reg_email:
                st.warning("请填写所有字段")
            elif reg_pass != reg_pass2:
                st.warning("两次密码输入不一致")
            else:
                ok, msg = register_user(reg_user, reg_pass, reg_email)
                if ok:
                    st.success(f"✅ {msg}，请登录")
                else:
                    st.error(msg)
    
    with auth_tab[2]:
        st.markdown('<div class="login-title">重置密码</div>', unsafe_allow_html=True)
        reset_user = st.text_input("用户名", placeholder="注册时的用户名", key="reset_user")
        reset_email = st.text_input("注册邮箱", placeholder="注册时填写的邮箱", key="reset_email")
        reset_new = st.text_input("新密码", type="password", placeholder="6个字符以上", key="reset_new")
        reset_new2 = st.text_input("确认新密码", type="password", placeholder="再次输入新密码", key="reset_new2")
        if st.button("🔄 重置密码", type="primary", use_container_width=True):
            if not reset_user or not reset_email or not reset_new:
                st.warning("请填写所有字段")
            elif reset_new != reset_new2:
                st.warning("两次密码输入不一致")
            else:
                ok, msg = reset_password(reset_user, reset_email, reset_new)
                if ok:
                    st.success(f"✅ {msg}，请用新密码登录")
                else:
                    st.error(msg)
    
    st.markdown("""</div>""", unsafe_allow_html=True)
    
    st.markdown("""
    <div class="powered-by">
        🔥 Powered by 涅 · 数智质控 v0.6
    </div>
    """, unsafe_allow_html=True)
    
    st.stop()


# ========== 已登录：主界面 ==========
st.markdown("""
<div class="main-title">
    <span class="logo-icon">📊</span>
    <div>
        <h1>数智质控</h1>
        <div class="subtitle">SMART QUALITY CONTROL PLATFORM</div>
    </div>
</div>
""", unsafe_allow_html=True)

# 侧边栏用户信息
current_user_data = load_user(st.session_state.current_user) if st.session_state.current_user else None
is_admin = current_user_data and current_user_data.get("role") == "admin"
role_label = "👑 管理员" if is_admin else "👤 用户"

st.sidebar.markdown(f"""
<div style="background:rgba(26,115,232,0.1); border:1px solid rgba(26,115,232,0.2); border-radius:10px; padding:12px; margin-bottom:12px; text-align:center;">
    <div style="font-size:18px;">{role_label}</div>
    <div style="color:#ffffff; font-weight:600; font-size:14px;">{st.session_state.current_user}</div>
</div>
""", unsafe_allow_html=True)

if st.sidebar.button("🚪 退出登录", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.session_state.login_time = None
    st.session_state.pending_2fa_user = None
    st.rerun()

page_options = [
    "🏠 首页",
    "📤 数据上传",
    "🔧 项目设置",
    "🧠 模型训练",
    "📊 相关性分析",
    "🔍 数据智能分析",
    "📋 异常报告",
    "📁 模型管理",
    "💬 功能反馈"
]
if is_admin:
    page_options.append("⚙️ 用户管理")

page = st.sidebar.radio("功能导航", page_options)

# 首页快捷导航跳转
if "nav_target" in st.session_state:
    page = st.session_state.pop("nav_target")

# 侧边栏底部信息
st.sidebar.markdown("---")
st.sidebar.markdown("""
<div style="text-align:center; color:#546e7a; font-size:11px;">
    <b>数智质控</b> v0.6<br>
    Smart Quality Control Platform<br>
    🔥 Powered by 涅
</div>
""", unsafe_allow_html=True)

# ==================== 页面0：首页 ====================
if page == "🏠 首页":
    # 顶部欢迎
    st.markdown(f"""
    <div style="text-align:center; padding:30px 0 20px 0;">
        <div style="font-size:48px; margin-bottom:10px;">👋</div>
        <div style="font-size:24px; font-weight:700; color:#ffffff;">欢迎回来，{st.session_state.current_user}</div>
        <div style="font-size:14px; color:#78909c; margin-top:8px;">检测数据智能分析与质控预警系统</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 快速入口卡片
    st.subheader("🚀 快速开始")
    st.markdown("""
    <style>
        .quick-card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 14px;
            padding: 20px;
            cursor: pointer;
            transition: all 0.3s ease;
            text-align: center;
        }
        .quick-card:hover {
            border-color: var(--primary);
            box-shadow: 0 4px 16px rgba(26, 115, 232, 0.15);
        }
        .quick-icon {
            font-size: 32px;
            margin-bottom: 8px;
        }
        .quick-title {
            font-size: 15px;
            font-weight: 600;
            color: #ffffff;
            margin-bottom: 4px;
        }
        .quick-desc {
            font-size: 12px;
            color: #78909c;
        }
    </style>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button("📤\n上传数据\n导入Excel历史数据", use_container_width=True, key="nav_upload"):
            st.session_state["nav_target"] = "📤 数据上传"
            st.rerun()
    with col2:
        if st.button("🧠\n训练模型\n按分类智能建模", use_container_width=True, key="nav_train"):
            st.session_state["nav_target"] = "🧠 模型训练"
            st.rerun()
    with col3:
        if st.button("🔍\n智能分析\n检测数据异常", use_container_width=True, key="nav_analyze"):
            st.session_state["nav_target"] = "🔍 数据智能分析"
            st.rerun()
    with col4:
        if st.button("📋\n异常报告\n查看分析结果", use_container_width=True, key="nav_report"):
            st.session_state["nav_target"] = "📋 异常报告"
            st.rerun()
    
    st.markdown("---")
    
    # 系统状态
    st.subheader("📊 系统状态")
    saved = get_saved_models(st.session_state.current_user)
    model_count = len(saved)
    total_samples = sum(sm.get("sample_count", 0) for sm in saved) if saved else 0
    has_data = st.session_state.df_raw is not None
    data_rows = len(st.session_state.df_raw) if has_data else 0
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.metric("已保存模型", f"{model_count} 个")
    with col_b:
        st.metric("历史数据", f"{data_rows} 条" if has_data else "未上传")
    with col_c:
        st.metric("模型样本量", f"{total_samples} 条" if total_samples > 0 else "暂无")
    
    st.markdown("---")
    
    # 工作流程
    st.subheader("📋 使用流程")
    steps = [
        ("📤", "上传历史数据", "上传Excel文件，系统自动识别数值列和分类列"),
        ("🔧", "项目设置", "配置小数位等参数，选择分类字段"),
        ("🧠", "训练模型", "按分类分组训练，建立各项目的统计基线"),
        ("🔍", "数据智能分析", "上传新数据或手动输入，自动检测异常"),
        ("📋", "查看异常报告", "总览、项目汇总、明细筛选，一键导出"),
    ]
    for i, (icon, title, desc) in enumerate(steps):
        col_num, col_content = st.columns([1, 9])
        with col_num:
            st.markdown(f"<div style='font-size:24px; text-align:center;'>{icon}</div>", unsafe_allow_html=True)
        with col_content:
            st.markdown(f"**第{i+1}步：{title}** — {desc}")
        if i < len(steps) - 1:
            st.markdown("<div style='text-align:center; color:#37474f;'>↓</div>", unsafe_allow_html=True)


# ==================== 页面1：数据上传 ====================
if page == "📤 数据上传":
    st.header("📤 第一步：上传历史数据")
    
    st.warning("⚠️ **数据上传须知**\n\n"
               "- 请**剔除敏感信息**（客户名称、价格等），只保留检测项目和分类信息（如供应商、批次、产线等）\n"
               "- 数据格式：**列=检测项目**，**行=样品记录**（如果反过来系统会自动识别）\n"
               "- 保留分类列（如供应商、批次、类别），系统会按分类分别建模\n"
               "- 支持 .xlsx / .xls 格式，可一次上传多个文件")
    
    uploaded_files = st.file_uploader(
        "上传历史数据（支持多个文件）", 
        type=["xlsx", "xls"], 
        key="history_upload", 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        dfs = []
        for f in uploaded_files:
            try:
                df_temp = load_data(f)
                dfs.append(df_temp)
            except Exception as e:
                st.error(f"读取文件 {f.name} 失败：{e}")
        
        if dfs:
            if len(dfs) == 1:
                df_raw = dfs[0]
            else:
                df_raw = pd.concat(dfs, ignore_index=True)
                df_raw = clean_columns(df_raw)
            
            # 检测是否需要转置
            df_raw, was_transposed = transpose_if_needed(df_raw)
            if was_transposed:
                st.info("🔄 检测到数据方向为「行=项目」，已自动转置为标准格式")
            
            # 识别列
            numeric_cols = detect_numeric_cols(df_raw)
            candidates = get_classification_candidates(df_raw, numeric_cols)

            st.session_state.df_raw = df_raw
            st.session_state.numeric_cols = numeric_cols

            st.success(f"✅ 数据读取成功！共 {len(df_raw)} 条记录，识别到 {len(numeric_cols)} 个检测项目")

            # 数据预览
            with st.expander("📋 数据预览（前20行）", expanded=True):
                st.dataframe(df_raw.head(20), use_container_width=True)

            # 识别结果 + 分类列自选
            col1, col2 = st.columns(2)
            with col1:
                st.subheader(f"检测项目（{len(numeric_cols)}个）")
                for c in numeric_cols:
                    st.markdown(f"- {c}")
            with col2:
                if candidates:
                    st.subheader("📌 请勾选分类列")
                    st.caption("系统已预选推荐项，你可以根据需要调整。分类列将用于分组建模。")
                    selected_cats = []
                    for col_name, recommended in candidates.items():
                        unique_vals = df_raw[col_name].dropna().unique()
                        val_preview = "、".join([str(v) for v in unique_vals[:5]])
                        more = f"…等{len(unique_vals)}个值" if len(unique_vals) > 5 else f"（{len(unique_vals)}个值）"
                        label = f"**{col_name}**  示例：{val_preview}{more}"
                        if st.checkbox(label, value=recommended, key=f"sel_cat_{col_name}"):
                            selected_cats.append(col_name)
                    st.session_state.category_cols = selected_cats
                    if selected_cats:
                        st.success(f"已选择 {len(selected_cats)} 个分类列：{'、'.join(selected_cats)}")
                    else:
                        st.info("未选择分类列，将使用全部数据一起建模")
                else:
                    st.subheader("分类列")
                    st.info("未找到候选分类列，将使用全部数据建模")
                    st.session_state.category_cols = []
            
            # 如果没有识别到数值列
            if len(numeric_cols) == 0:
                st.error("❌ 未识别到数值列！可能原因：\n"
                        "1. 数据存储为文本格式，请检查Excel中数字列的格式\n"
                        "2. 列名中包含特殊字符\n"
                        "3. 数据确实不含数值\n\n"
                        "建议：在Excel中选中数值列 → 设置单元格格式为「数值」")
        
        st.markdown("---")
        st.markdown("👉 识别结果没问题？去 **🔧 项目设置** 配置小数位，然后去 **🧠 模型训练** 建模")


# ==================== 页面2：项目设置 ====================
elif page == "🔧 项目设置":
    st.header("🔧 项目设置")
    
    df_raw = st.session_state.get("df_raw")
    numeric_cols = st.session_state.get("numeric_cols", [])
    
    if df_raw is None:
        st.warning("⚠️ 请先在「数据上传」页面上传数据")
    elif len(numeric_cols) == 0:
        st.error("❌ 未识别到数值列，请检查数据格式")
    else:
        st.markdown("为每个检测项目设置小数位数，这会影响模型精度和显示格式。")
        
        # 初始化小数位设置
        if "decimal_settings" not in st.session_state:
            st.session_state.decimal_settings = {}
        
        # 自动推断小数位
        auto_decimals = {}
        for col in numeric_cols:
            series = pd.to_numeric(df_raw[col], errors='coerce').dropna()
            if len(series) > 0:
                # 取中位数的小数位作为默认值
                sample_vals = series.head(100)
                max_dec = 0
                for v in sample_vals:
                    s = str(v)
                    if '.' in s:
                        d = len(s.split('.')[1])
                        max_dec = max(max_dec, d)
                auto_decimals[col] = min(max_dec, 6)
            else:
                auto_decimals[col] = 2
        
        # 显示设置表格
        st.subheader("项目选取与小数位设置")
        st.caption("勾选「纳入」的项目将用于模型训练。小数位影响模型精度和显示格式。")

        # 初始化项目选择
        if "selected_cols" not in st.session_state:
            st.session_state.selected_cols = list(numeric_cols)

        cols_per_row = 3
        for i in range(0, len(numeric_cols), cols_per_row):
            row_cols = st.columns(cols_per_row)
            for j, col_name in enumerate(numeric_cols[i:i+cols_per_row]):
                with row_cols[j]:
                    st.markdown(f"**{col_name}**")
                    use_col = st.checkbox(
                        "纳入模型",
                        value=col_name in st.session_state.selected_cols,
                        key=f"use_{col_name}",
                    )
                    default_dec = st.session_state.decimal_settings.get(col_name, auto_decimals.get(col_name, 2))
                    dec = st.number_input(
                        "小数位",
                        min_value=0, max_value=6,
                        value=default_dec,
                        key=f"dec_{col_name}",
                        help=f"自动推断：{auto_decimals.get(col_name, 2)}位"
                    )
                    st.session_state.decimal_settings[col_name] = int(dec)

        # 更新已选列表
        st.session_state.selected_cols = [
            c for c in numeric_cols if st.session_state.get(f"use_{c}", True)
        ]
        st.markdown(f"已选 **{len(st.session_state.selected_cols)}/{len(numeric_cols)}** 个项目用于建模")
        
        st.markdown("---")
        st.info("💡 设置完成后，去 **🧠 模型训练** 选择分类方式并开始建模")


# ==================== 页面3：模型训练 ====================
elif page == "🔧 项目设置" or page == "🧠 模型训练":
    if page == "🧠 模型训练":
        st.header("🧠 第三步：选择分类方式并训练模型")
        
        df_raw = st.session_state.get("df_raw")
        numeric_cols = st.session_state.get("numeric_cols", [])
        category_cols = st.session_state.get("category_cols", [])
        decimal_settings = st.session_state.get("decimal_settings", {})

        # 应用项目选取
        selected_cols = st.session_state.get("selected_cols", numeric_cols)
        numeric_cols = [c for c in numeric_cols if c in selected_cols]

        if df_raw is None or len(numeric_cols) == 0:
            st.warning("⚠️ 请先上传数据并确认项目设置（可能所有项目都被取消了选择）")
        else:
            # 选择分类方式
            st.subheader("选择建模分类")
            
            if category_cols:
                st.markdown("选择按哪个分类列分组建模（建议选择，不同分组的数据特征差异大）：")
                
                # 勾选分类列
                selected_group_col = st.selectbox(
                    "分类依据",
                    ["不分组（使用全部数据）"] + category_cols,
                    index=1 if category_cols else 0
                )
                
                if selected_group_col != "不分组（使用全部数据）":
                    group_values = df_raw[selected_group_col].dropna().unique().tolist()
                    
                    # 显示各类别数据量
                    st.markdown(f"**{selected_group_col}** 包含以下类别：")
                    for gv in group_values:
                        count = len(df_raw[df_raw[selected_group_col] == gv])
                        st.markdown(f"- {gv}：**{count}** 条")
                    
                    # 选择要建模的类别
                    selected_groups = st.multiselect(
                        "选择要建模的类别（不选则全部建模）",
                        group_values,
                        default=group_values
                    )
                else:
                    selected_group_col = None
                    selected_groups = [None]
            else:
                st.info("未识别到分类列，将使用全部数据一起建模")
                selected_group_col = None
                selected_groups = [None]
            
            # 数据量控制
            st.subheader("📏 数据量控制")
            st.markdown("设置单个模型保留的最大数据量。超过上限时自动淘汰最旧数据，保持模型基于最新数据训练。")
            
            max_count = st.slider(
                "单个模型最大数据量",
                min_value=100, max_value=2000, value=st.session_state.max_data_count,
                step=100,
                help="建议500-2000条。2000条以内运算秒出，不会慢。"
            )
            st.session_state.max_data_count = max_count
            
            # 预览各分类数据量是否超限
            if selected_group_col:
                for gv in (selected_groups or []):
                    if gv is not None:
                        count = len(df_raw[df_raw[selected_group_col] == gv])
                        if count > max_count:
                            st.warning(f"⚠️ 「{gv}」有 {count} 条数据，超过上限 {max_count}，训练时将自动使用最近 {max_count} 条")
            else:
                total = len(df_raw)
                if total > max_count:
                    st.warning(f"⚠️ 全部数据共 {total} 条，超过上限 {max_count}，训练时将自动使用最近 {max_count} 条")
                else:
                    st.success(f"✅ 当前数据量 {total} 条，在上限范围内")
            
            # 训练按钮
            if st.button("🚀 开始训练", type="primary"):
                progress = st.progress(0)
                trained_models = []
                
                df_numeric = df_raw[numeric_cols].apply(pd.to_numeric, errors='coerce')
                
                for idx, group_val in enumerate(selected_groups):
                    if group_val is not None:
                        mask = df_raw[selected_group_col] == group_val
                        df_num_filtered = df_numeric[mask]
                    else:
                        df_num_filtered = df_numeric
                    
                    if len(df_num_filtered) < 5:
                        st.warning(f"⚠️ {group_val or '全部'}数据不足5条，跳过")
                        continue
                    
                    model = build_model(
                        df_num_filtered, numeric_cols,
                        group_col=selected_group_col,
                        group_value=group_val,
                        decimal_settings=decimal_settings,
                        max_data_count=st.session_state.get("max_data_count", 500)
                    )
                    
                    # 保存模型
                    model_name = f"{selected_group_col}_{group_val}" if group_val else "全部数据"
                    save_model(model, model_name, st.session_state.current_user)
                    st.session_state.models[model_name] = model
                    trained_models.append((model_name, model))
                    
                    progress.progress((idx + 1) / len(selected_groups))
                
                st.success(f"✅ 训练完成！共训练 {len(trained_models)} 个模型")
                
                # 数据量提示
                max_count = st.session_state.get("max_data_count", 500)
                for model_name, model in trained_models:
                    if model.get("data_trimmed"):
                        st.warning(f"💡 模型「{model_name}」数据量超过上限（{max_count}条），已自动使用最近 {max_count} 条数据训练")
                
                # 展示结果
                for model_name, model in trained_models:
                    with st.expander(f"📊 {model_name}（{model['sample_count']}条数据）", expanded=True):
                        # 统计特征
                        stats_data = []
                        for col, info in model["item_stats"].items():
                            stats_data.append({
                                "项目": col,
                                "样本量": info['count'],
                                "均值": info['mean'],
                                "标准差": info['std'],
                                "正常范围(IQR)": f"{info['lower_fence']} ~ {info['upper_fence']}"
                            })
                        st.dataframe(pd.DataFrame(stats_data), use_container_width=True, hide_index=True)
                        
                        # 相关性
                        if model.get("strong_pairs"):
                            st.markdown(f"**强相关项目对（{len(model['strong_pairs'])}对）：**")
                            for pair in model["strong_pairs"]:
                                icon = "📈" if pair["direction"] == "正相关" else "📉"
                                st.markdown(f"{icon} **{pair['col_a']}** ↔ **{pair['col_b']}**：{pair['direction']}，r={pair['r']}")
                        else:
                            st.warning("未发现强相关项目对，可能数据量不足或分组不够细")
                
                st.markdown("---")
                st.markdown("👉 去往 **📊 相关性分析** 查看详细相关性，或 **🔍 数据智能分析** 开始使用")


# ==================== 页面4：相关性分析 ====================
elif page == "📊 相关性分析":
    st.header("📊 相关性分析")
    st.markdown("分析检测项目之间的统计相关性，发现项目间的内在规律。")
    
    # 选择模型
    saved = get_saved_models(st.session_state.current_user)
    current_models = list(st.session_state.models.keys())
    all_model_names = list(set(sm["name"] for sm in saved) | set(current_models))
    
    if not all_model_names:
        st.warning("⚠️ 请先训练模型")
    else:
        # 用友好名称做选项映射
        name_map = {}
        for mn in all_model_names:
            m = st.session_state.models.get(mn)
            display = get_model_display_name(mn, m, st.session_state.current_user)
            if m:
                group_info = f"（{m.get('group_value', '')}）" if m.get('group_value') and m.get('group_value') != '全部' else ""
                display = f"{display}{group_info} | {m.get('sample_count', '?')}条数据"
            else:
                for sm in saved:
                    if sm["name"] == mn:
                        group_info = f"（{sm['group']}）" if sm['group'] != '全部' else ""
                        display = f"{sm['display_name']}{group_info} | {sm['sample_count']}条数据"
                        break
            name_map[display] = mn
        
        selected_display = st.selectbox("选择模型", list(name_map.keys()), key="corr_model_select")
        selected_model_name = name_map[selected_display]
        
        # 加载模型
        model = st.session_state.models.get(selected_model_name)
        if not model:
            for sm in saved:
                if sm["name"] == selected_model_name:
                    model = load_model_file(sm["filepath"])
                    st.session_state.models[selected_model_name] = model
                    break
        
        if model:
            group_info = f"（{model.get('group_col', '')}：{model.get('group_value', '')}）" if model.get('group_value') and model.get('group_value') != '全部' else ""
            st.caption(f"模型：{selected_model_name}{group_info} | 样本量：{model['sample_count']}")
            
            correlations = model.get("correlations", {})
            strong_pairs = model.get("strong_pairs", [])
            
            if not correlations:
                st.warning("该模型暂无相关性数据")
            else:
                # ===== 1. 强相关项目对列表 =====
                st.subheader("🔗 强相关项目对")
                
                if strong_pairs:
                    st.markdown(f"以下项目对之间存在**显著相关关系**（|r| > 0.5），共 **{len(strong_pairs)}** 对：")
                    
                    pair_data = []
                    for pair in strong_pairs:
                        if pair["r"] > 0.8:
                            strength = "🔴 极强"
                        elif pair["r"] > 0.6 or pair["r"] < -0.8:
                            strength = "🟠 强"
                        else:
                            strength = "🟡 中等"
                        
                        pair_data.append({
                            "项目A": pair["col_a"],
                            "项目B": pair["col_b"],
                            "变化关系": "同步变化" if pair["direction"] == "正相关" else "此消彼长",
                            "相关度": pair["r"],
                            "强度": strength,
                            "通俗理解": f"{pair['col_a']}涨 → {pair['col_b']}也涨" if pair["direction"] == "正相关" else f"{pair['col_a']}涨 → {pair['col_b']}降"
                        })
                    
                    pair_df = pd.DataFrame(pair_data)
                    st.dataframe(pair_df, use_container_width=True, hide_index=True)
                    
                    # 简明解读
                    st.subheader("💡 规律解读")
                    for pair in strong_pairs[:10]:  # 最多展示10对
                        icon = "📈" if pair["direction"] == "正相关" else "📉"
                        st.markdown(f"{icon} **{pair['col_a']}** 与 **{pair['col_b']}**：{pair['direction']}（相关度 = {pair['r']}）")
                        if pair["direction"] == "负相关":
                            st.caption(f"→ {pair['col_a']}越高，{pair['col_b']}越低，两者反着走")
                        else:
                            st.caption(f"→ {pair['col_a']}和{pair['col_b']}同涨同跌，一个高了另一个也高")
                else:
                    st.info("未发现强相关项目对（|r| > 0.5）。可能原因：数据量不足、混合了不同类型数据、或项目间确实无强相关。")
                
                # ===== 2. 相关性热力图 =====
                st.subheader("🌡️ 相关性热力图")
                
                corr_df = pd.DataFrame(correlations)
                fig_heatmap = px.imshow(
                    corr_df, 
                    text_auto=".2f", 
                    aspect="auto",
                    color_continuous_scale="RdBu_r", 
                    range_color=[-1, 1],
                    title="项目间相关性热力图（红色=正相关，蓝色=负相关）"
                )
                fig_heatmap.update_layout(
                    height=max(500, len(corr_df.columns) * 45 + 100),
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font_color='#b0bec5',
                    title_font_color='#4fc3f7'
                )
                st.plotly_chart(fig_heatmap, use_container_width=True)
                
                st.caption("💡 **读图说明**：越红→同涨同跌关系越强，越蓝→反着走关系越强，白色→没啥关系。")
                
                # ===== 3. 散点图 =====
                if strong_pairs:
                    st.subheader("📈 散点图")
                    st.markdown("选择一对项目，查看数据分布和回归趋势线：")
                    
                    pair_options = [f"{p['col_a']} ↔ {p['col_b']} (r={p['r']})" for p in strong_pairs]
                    selected_pair_idx = st.selectbox("选择项目对", range(len(pair_options)), format_func=lambda i: pair_options[i], key="scatter_pair")
                    
                    selected_pair = strong_pairs[selected_pair_idx]
                    col_a, col_b = selected_pair["col_a"], selected_pair["col_b"]
                    
                    # 获取原始数据
                    df_raw = st.session_state.get("df_raw")
                    if df_raw is not None:
                        numeric_cols_session = st.session_state.get("numeric_cols", [])
                        df_numeric = df_raw[numeric_cols_session].apply(pd.to_numeric, errors='coerce')
                        
                        # 如果模型是分组的，过滤数据
                        if model.get("group_col") and model.get("group_value"):
                            mask = df_raw[model["group_col"]] == model["group_value"]
                            df_numeric = df_numeric[mask]
                        
                        mask_valid = df_numeric[col_a].notna() & df_numeric[col_b].notna()
                        plot_data = df_numeric.loc[mask_valid, [col_a, col_b]]
                        
                        if len(plot_data) > 0:
                            fig_scatter = px.scatter(
                                plot_data, x=col_a, y=col_b,
                                title=f"{col_a} 与 {col_b} 散点图（r = {selected_pair['r']}）",
                                opacity=0.6,
                                labels={col_a: col_a, col_b: col_b}
                            )
                            
                            # 添加回归线
                            x_vals = plot_data[col_a].values
                            y_vals = plot_data[col_b].values
                            slope, intercept, r_val, p_val, std_err = stats.linregress(x_vals, y_vals)
                            x_range = np.linspace(x_vals.min(), x_vals.max(), 100)
                            y_pred = slope * x_range + intercept
                            
                            fig_scatter.add_trace(go.Scatter(
                                x=x_range, y=y_pred,
                                mode='lines', name='回归线',
                                line=dict(color='red', dash='dash')
                            ))
                            
                            # 添加置信区间
                            residual_std = (y_vals - (slope * x_vals + intercept)).std()
                            fig_scatter.add_trace(go.Scatter(
                                x=np.concatenate([x_range, x_range[::-1]]),
                                y=np.concatenate([y_pred + 2.5 * residual_std, (y_pred - 2.5 * residual_std)[::-1]]),
                                fill='toself', fillcolor='rgba(0,100,0,0.1)',
                                line=dict(color='rgba(0,100,0,0)'), name='95%置信区间'
                            ))
                            
                            fig_scatter.update_layout(height=500, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color='#b0bec5', title_font_color='#4fc3f7')
                            st.plotly_chart(fig_scatter, use_container_width=True)
                            
                            st.caption(f"回归方程：{col_b} = {slope:.4f} × {col_a} + {intercept:.4f} | R² = {r_val**2:.4f}")
                
                # ===== 4. 各项目数据分布 =====
                with st.expander("📊 各项目数据分布图", expanded=False):
                    df_raw = st.session_state.get("df_raw")
                    if df_raw is not None:
                        numeric_cols_session = st.session_state.get("numeric_cols", [])
                        df_numeric = df_raw[numeric_cols_session].apply(pd.to_numeric, errors='coerce')
                        
                        if model.get("group_col") and model.get("group_value"):
                            mask = df_raw[model["group_col"]] == model["group_value"]
                            df_numeric = df_numeric[mask]
                        
                        for col in numeric_cols_session:
                            series = df_numeric[col].dropna()
                            if len(series) < 3:
                                continue
                            stats_info = model["item_stats"].get(col, {})
                            fig = go.Figure()
                            fig.add_trace(go.Histogram(x=series, name="数据分布", nbinsx=30, marker_color='steelblue'))
                            if stats_info:
                                fig.add_vline(x=stats_info["mean"], line_dash="dash", line_color="green", annotation_text=f"均值:{stats_info['mean']}")
                                fig.add_vrect(x0=stats_info["lower_fence"], x1=stats_info["upper_fence"], 
                                             fillcolor="green", opacity=0.1, annotation_text="正常范围")
                            fig.update_layout(title=f"{col} 分布", showlegend=False, height=300, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color='#b0bec5', title_font_color='#4fc3f7')
                            st.plotly_chart(fig, use_container_width=True)


# ==================== 页面5：数据智能分析 ====================
elif page == "🔍 数据智能分析":
    st.header("🔍 数据智能分析")
    
    # 获取所有可用模型
    saved = get_saved_models(st.session_state.current_user)
    current_models = list(st.session_state.models.keys())
    all_model_names = list(set(sm["name"] for sm in saved) | set(current_models))
    
    if not all_model_names:
        st.warning("⚠️ 请先训练模型")
    else:
        # 输入方式决定流程
        input_method = st.radio("数据输入方式", ["手动输入", "上传Excel"], horizontal=True)
        
        if input_method == "手动输入":
            # 手动输入：先选模型，再填数据
            # 用友好名称做选项映射
            name_map = {}
            for mn in all_model_names:
                m = st.session_state.models.get(mn)
                display = get_model_display_name(mn, m, st.session_state.current_user)
                if m:
                    group_info = f"（{m.get('group_value', '')}）" if m.get('group_value') and m.get('group_value') != '全部' else ""
                    display = f"{display}{group_info} | {m.get('sample_count', '?')}条数据"
                else:
                    for sm in saved:
                        if sm["name"] == mn:
                            group_info = f"（{sm['group']}）" if sm['group'] != '全部' else ""
                            display = f"{sm['display_name']}{group_info} | {sm['sample_count']}条数据"
                            break
                name_map[display] = mn
            
            selected_display = st.selectbox("选择模型", list(name_map.keys()))
            selected_model_name = name_map[selected_display]
            
            # 加载模型
            model = st.session_state.models.get(selected_model_name)
            if not model:
                for sm in saved:
                    if sm["name"] == selected_model_name:
                        model = load_model_file(sm["filepath"])
                        st.session_state.models[selected_model_name] = model
                        break
            
            if model:
                group_info = f"（{model.get('group_col', '')}：{model.get('group_value', '全部')}）" if model.get('group_value') != '全部' else ""
                st.info(f"当前模型{group_info}，样本量：{model['sample_count']}，训练时间：{model['trained_at']}")
                
                new_data = {}
                numeric_cols = model.get("numeric_cols", list(model.get("item_stats", {}).keys()))
                
                cols_per_row = 3
                for i in range(0, len(numeric_cols), cols_per_row):
                    row_cols = st.columns(cols_per_row)
                    for j, col_name in enumerate(numeric_cols[i:i+cols_per_row]):
                        with row_cols[j]:
                            stats_info = model["item_stats"].get(col_name, {})
                            hint = f"均值:{stats_info.get('mean','?')} 范围:{stats_info.get('lower_fence','?')}~{stats_info.get('upper_fence','?')}" if stats_info else ""
                            dec = stats_info.get("decimals", 3)
                            fmt = f"%.{dec}f"
                            val = st.number_input(f"{col_name}", key=f"input_{col_name}", help=hint, format=fmt)
                            new_data[col_name] = val
                
                if st.button("🔎 开始分析", type="primary"):
                    if not new_data:
                        st.warning("请输入数据")
                    else:
                        anomalies = check_anomaly(new_data, model)
                        # 统一格式，给手动输入的异常也加行号和模型信息
                        for a in anomalies:
                            a["行号"] = 1
                            a["使用模型"] = get_model_display_name(selected_model_name, model, st.session_state.current_user)
                        st.session_state["latest_anomalies"] = anomalies
                        st.session_state["latest_data"] = {"mode": "manual", "data": new_data}
                        st.session_state["latest_model"] = get_model_display_name(selected_model_name, model, st.session_state.current_user)
                        # 手动输入的数据自动记录到新数据中，不丢失
                        new_row_df = pd.DataFrame([new_data])
                        if st.session_state.new_data is None:
                            st.session_state.new_data = new_row_df.copy()
                        else:
                            st.session_state.new_data = pd.concat(
                                [st.session_state.new_data, new_row_df], ignore_index=True
                            )
                        
                        if anomalies:
                            st.error(f"⚠️ 发现 {len(anomalies)} 个异常！")
                            anomaly_df = pd.DataFrame(anomalies)
                            st.dataframe(anomaly_df, use_container_width=True, hide_index=True)
                        else:
                            st.success("✅ 所有检测项目均在正常范围内")
                        
                        # 滚动更新
                        st.markdown("---")
                        st.subheader("🔄 模型滚动更新")
                        if len(anomalies) > 0:
                            st.warning("⚠️ 当前数据存在异常，建议确认数据无误后再纳入模型。")
                        
                        col_a, col_b = st.columns(2)
                        with col_a:
                            if st.button("✅ 确认数据无误，纳入模型", type="primary", key="update_model_yes"):
                                df_raw = st.session_state.get("df_raw")
                                if df_raw is not None:
                                    new_row_df = pd.DataFrame([new_data])
                                    df_raw = pd.concat([df_raw, new_row_df], ignore_index=True)
                                    st.session_state.df_raw = df_raw
                                    # 注：new_data 已在「开始分析」时记录，这里不再重复追加

                                    numeric_cols_session = st.session_state.get("numeric_cols", [])
                                    selected_cols = st.session_state.get("selected_cols", numeric_cols_session)
                                    numeric_cols_session = [c for c in numeric_cols_session if c in selected_cols]
                                    decimal_settings = st.session_state.get("decimal_settings", {})
                                    df_numeric = df_raw[numeric_cols_session].apply(pd.to_numeric, errors='coerce')
                                    
                                    group_col = model.get("group_col")
                                    group_value = model.get("group_value")
                                    
                                    if group_col and group_value:
                                        mask = df_raw[group_col] == group_value
                                        df_num_filtered = df_numeric[mask]
                                    else:
                                        df_num_filtered = df_numeric
                                    
                                    new_model = build_model(
                                        df_num_filtered, numeric_cols_session,
                                        group_col=group_col,
                                        group_value=group_value,
                                        decimal_settings=decimal_settings,
                                        max_data_count=st.session_state.get("max_data_count", 500)
                                    )

                                    save_model(new_model, selected_model_name, st.session_state.current_user)
                                    st.session_state.models[selected_model_name] = new_model

                                    # 从 new_data 中移除已纳入模型的数据
                                    if st.session_state.new_data is not None and len(st.session_state.new_data) > 0:
                                        st.session_state.new_data = st.session_state.new_data.iloc[:-1]
                                        if len(st.session_state.new_data) == 0:
                                            st.session_state.new_data = None

                                    st.success(f"✅ 模型已更新！样本量：{model['sample_count']} → {new_model['sample_count']}")
                                    st.rerun()
                                else:
                                    st.error("无法更新模型：历史数据不可用")
                        
                        with col_b:
                            if st.button("❌ 不纳入模型", key="update_model_no"):
                                st.info("数据未纳入模型，模型保持不变")
        
        else:
            # 上传Excel：自动匹配模型
            st.markdown("💡 上传Excel后，系统将根据分类列（如供应商、批次、产线等）**自动匹配对应模型**进行分析。")
            
            daily_file = st.file_uploader("上传检测数据", type=["xlsx", "xls"], key="daily_upload")
            
            if daily_file:
                df_daily = load_data(daily_file)
                st.dataframe(df_daily.head(20), use_container_width=True)
                
                # 识别分类列（用户自选，系统推荐）
                daily_numeric_cols = detect_numeric_cols(df_daily)
                daily_candidates = get_classification_candidates(df_daily, daily_numeric_cols)
                if daily_candidates:
                    st.subheader("📌 请勾选用于匹配模型的分类列")
                    st.caption("系统会按此列的值自动匹配对应模型进行异常检测。")
                    daily_category_cols = []
                    for col_name, recommended in daily_candidates.items():
                        unique_vals = df_daily[col_name].dropna().unique()
                        val_preview = "、".join([str(v) for v in unique_vals[:5]])
                        more = f"…等{len(unique_vals)}个值" if len(unique_vals) > 5 else f"（{len(unique_vals)}个值）"
                        label = f"**{col_name}**  示例：{val_preview}{more}"
                        if st.checkbox(label, value=recommended, key=f"daily_cat_{col_name}"):
                            daily_category_cols.append(col_name)
                else:
                    daily_category_cols = []
                
                # 构建所有可用模型的映射：group_col + group_value → model
                model_map = {}  # key: (group_col, group_value), value: model
                no_group_models = []  # 没有分组的模型
                
                for mn in all_model_names:
                    m = st.session_state.models.get(mn)
                    if not m:
                        for sm in saved:
                            if sm["name"] == mn:
                                m = load_model_file(sm["filepath"])
                                st.session_state.models[mn] = m
                                break
                    if m:
                        gcol = m.get("group_col")
                        gval = m.get("group_value")
                        if gcol and gval and gval != "全部":
                            model_map[(gcol, str(gval))] = m
                        else:
                            no_group_models.append(m)
                
                # 对Excel中每条数据，自动匹配模型并分析
                if st.button("🔎 开始批量分析", type="primary"):
                    all_results = []
                    all_anomaly_details = []  # 存每条记录的异常详情
                    total_rows = len(df_daily)
                    
                    for idx in range(total_rows):
                        row = df_daily.iloc[idx]
                        row_numeric = {col: pd.to_numeric(row.get(col), errors='coerce') for col in daily_numeric_cols}
                        
                        # 自动匹配模型
                        matched_model = None
                        matched_model_name = None
                        
                        # 先按分类列匹配
                        for cat_col in daily_category_cols:
                            cat_val = str(row.get(cat_col, ""))
                            if cat_val and cat_val != "nan":
                                # 在模型映射中查找
                                for (gcol, gval), m in model_map.items():
                                    if cat_val == gval:
                                        matched_model = m
                                        matched_model_name = get_model_display_name(gcol + "_" + gval, m, st.session_state.current_user)
                                        break
                            if matched_model:
                                break
                        
                        # 没匹配到分组模型，用通用模型
                        if not matched_model and no_group_models:
                            matched_model = no_group_models[0]
                            matched_model_name = get_model_display_name("全部数据", matched_model, st.session_state.current_user)

                        if matched_model:
                            anomalies = check_anomaly(row_numeric, matched_model)
                            desc_list = []
                            for a in anomalies:
                                item = a["项目"]
                                atype = a["类型"]
                                # 逻辑异常已经包含项目名，不需要重复
                                if "关系异常" in atype:
                                    desc_list.append(atype)
                                else:
                                    desc_list.append(f"{item}{atype}")
                            result = {
                                "行号": idx + 1,
                                "使用模型": matched_model_name,
                                "异常说明": "、".join(desc_list) if desc_list else "正常",
                            }
                            all_results.append(result)
                            # 保存异常详情，供异常报告页使用
                            if anomalies:
                                for a in anomalies:
                                    a_copy = a.copy()
                                    a_copy["行号"] = idx + 1
                                    a_copy["使用模型"] = matched_model_name
                                    all_anomaly_details.append(a_copy)
                        else:
                            all_results.append({
                                "行号": idx + 1,
                                "使用模型": "⚠️ 未匹配",
                                "异常说明": "未找到匹配模型"
                            })
                    
                    # 保存到session_state，供异常报告页使用
                    st.session_state["latest_anomalies"] = all_anomaly_details
                    st.session_state["latest_data"] = {"mode": "batch", "total_rows": total_rows}
                    st.session_state["latest_model"] = "批量分析"
                    
                    # 展示结果
                    result_df = pd.DataFrame(all_results)

                    # 统计
                    anomaly_rows = result_df[result_df["异常说明"] != "正常"]
                    normal_rows = result_df[result_df["异常说明"] == "正常"]
                    no_model_rows = result_df[result_df["使用模型"] == "⚠️ 未匹配"]
                    
                    st.subheader("📊 批量分析结果")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("总记录数", f"{total_rows} 条")
                    with col2:
                        st.metric("检测正常", f"{len(normal_rows)} 条")
                    with col3:
                        st.metric("发现异常", f"{len(anomaly_rows)} 条")
                    
                    if len(no_model_rows) > 0:
                        st.warning(f"⚠️ 有 {len(no_model_rows)} 条数据未匹配到模型，请检查分类列是否与训练时一致")
                    
                    # 展示结果表格
                    if len(anomaly_rows) > 0:
                        st.error(f"⚠️ 发现 {len(anomaly_rows)} 条异常记录：")
                        st.dataframe(anomaly_rows, use_container_width=True, hide_index=True)
                    
                    with st.expander("查看全部结果"):
                        st.dataframe(result_df, use_container_width=True, hide_index=True)
                    
                    # 滚动更新
                    st.markdown("---")
                    st.subheader("🔄 模型滚动更新")
                    st.markdown("将本次上传的数据纳入模型，保持模型最新。")
                    
                    col_a, col_b = st.columns(2)
                    with col_a:
                        if st.button("✅ 确认数据无误，全部纳入模型", type="primary", key="update_model_batch_yes"):
                            df_raw = st.session_state.get("df_raw")
                            if df_raw is not None:
                                # 追加所有数据
                                df_daily_numeric = df_daily[daily_numeric_cols].apply(pd.to_numeric, errors='coerce')
                                df_raw = pd.concat([df_raw, df_daily], ignore_index=True)
                                st.session_state.df_raw = df_raw
                                # 累积新数据用于智能诊断
                                if st.session_state.new_data is None:
                                    st.session_state.new_data = df_daily.copy()
                                else:
                                    st.session_state.new_data = pd.concat(
                                        [st.session_state.new_data, df_daily], ignore_index=True
                                    )
                                
                                # 重新训练所有受影响的模型
                                numeric_cols_session = st.session_state.get("numeric_cols", [])
                                selected_cols = st.session_state.get("selected_cols", numeric_cols_session)
                                numeric_cols_session = [c for c in numeric_cols_session if c in selected_cols]
                                decimal_settings = st.session_state.get("decimal_settings", {})
                                df_numeric = df_raw[numeric_cols_session].apply(pd.to_numeric, errors='coerce')
                                
                                updated_count = 0
                                for mn in all_model_names:
                                    m = st.session_state.models.get(mn)
                                    if m:
                                        group_col = m.get("group_col")
                                        group_value = m.get("group_value")
                                        if group_col and group_value:
                                            mask = df_raw[group_col] == group_value
                                            df_num_filtered = df_numeric[mask]
                                        else:
                                            df_num_filtered = df_numeric
                                        
                                        new_model = build_model(
                                            df_num_filtered, numeric_cols_session,
                                            group_col=group_col,
                                            group_value=group_value,
                                            decimal_settings=decimal_settings,
                                            max_data_count=st.session_state.get("max_data_count", 500)
                                        )
                                        save_model(new_model, mn, st.session_state.current_user)
                                        st.session_state.models[mn] = new_model
                                        updated_count += 1
                                
                                # 清除 new_data（这批数据已全部纳入模型）
                                st.session_state.new_data = None

                                st.success(f"✅ 已将 {total_rows} 条数据纳入模型，共更新 {updated_count} 个模型")
                                st.rerun()
                            else:
                                st.error("无法更新模型：历史数据不可用")
                    
                    with col_b:
                        if st.button("❌ 不纳入模型", key="update_model_batch_no"):
                            st.info("数据未纳入模型，模型保持不变")


# ==================== 页面6：异常报告 ====================
elif page == "📋 异常报告":
    st.header("📋 异常报告")
    
    anomalies = st.session_state.get("latest_anomalies", [])
    latest_data = st.session_state.get("latest_data", {})
    model_name = st.session_state.get("latest_model", "")
    
    if not anomalies and not latest_data:
        st.info("暂无检测记录，请先进行数据智能分析")
    elif not anomalies:
        st.success(f"✅ 最近一次检测（{model_name}）所有项目均在正常范围内")
    else:
        # ===== 第一层：总览仪表盘 =====
        st.subheader("📊 总览")
        
        # 统计各类异常数量
        severe_count = len([a for a in anomalies if "严重" in a.get("类型", "")])
        mild_count = len([a for a in anomalies if "🟡" in a.get("类型", "")])
        logic_count = len([a for a in anomalies if "关系异常" in a.get("类型", "")])
        extreme_count = len([a for a in anomalies if "创历史新高" in a.get("类型", "") or "创历史新低" in a.get("类型", "")])
        
        # 涉及的唯一项目数
        affected_items = set(a["项目"] for a in anomalies)
        affected_count = len(affected_items)
        
        # 涉及的唯一行号数（异常记录条数）
        affected_rows = set(a.get("行号", 0) for a in anomalies)
        affected_row_count = len(affected_rows)
        
        total_anomaly_count = len(anomalies)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("异常项目数", f"{affected_count}", delta=f"共{total_anomaly_count}项异常")
        with col2:
            st.metric("🔴 严重异常", f"{severe_count}", delta=None, delta_color="inverse")
        with col3:
            st.metric("🟡 轻度异常", f"{mild_count}", delta=None, delta_color="inverse")
        with col4:
            st.metric("🔵 逻辑异常", f"{logic_count}", delta=None, delta_color="inverse")
        
        col5, col6, col7, col8 = st.columns(4)
        with col5:
            st.metric("⬆️⬇️ 历史极值", f"{extreme_count}")
        with col6:
            st.metric("异常数据条数", f"{affected_row_count}")
        with col7:
            mode_info = latest_data.get("mode", "manual")
            st.metric("分析模式", "批量上传" if mode_info == "batch" else "手动输入")
        with col8:
            st.metric("使用模型", model_name if len(model_name) <= 15 else model_name[:15] + "…")
        
        st.markdown("---")
        

        # ===== 第四层：异常概览图 =====
        st.subheader("📊 异常概览图")

        # 直接从 anomalies 构建图表数据
        chart_items = sorted(set(a["项目"] for a in anomalies))
        chart_severe = [sum(1 for a in anomalies if a["项目"] == item and "严重" in a.get("类型", "")) for item in chart_items]
        chart_mild = [sum(1 for a in anomalies if a["项目"] == item and "🟡" in a.get("类型", "")) for item in chart_items]
        chart_logic = [sum(1 for a in anomalies if a["项目"] == item and "关系异常" in a.get("类型", "")) for item in chart_items]
        chart_extreme = [sum(1 for a in anomalies if a["项目"] == item and ("创历史新高" in a.get("类型", "") or "创历史新低" in a.get("类型", ""))) for item in chart_items]

        fig = go.Figure()
        fig.add_trace(go.Bar(name="严重异常", x=chart_items, y=chart_severe, marker_color="#d32f2f"))
        fig.add_trace(go.Bar(name="轻度异常", x=chart_items, y=chart_mild, marker_color="#ff9800"))
        fig.add_trace(go.Bar(name="逻辑异常", x=chart_items, y=chart_logic, marker_color="#1a73e8"))
        fig.add_trace(go.Bar(name="历史极值", x=chart_items, y=chart_extreme, marker_color="#607d8b"))

        fig.update_layout(
            barmode="stack",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font_color="#b0bec5",
            xaxis_title="检测项目",
            yaxis_title="异常次数",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            height=350,
            margin=dict(l=40, r=20, t=30, b=60),
        )
        fig.update_xaxes(tickangle=-30)
        st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        
        # ===== 第五层：可筛选的异常明细表 =====
        st.subheader("🔍 异常明细")
        st.markdown("筛选和查看每条异常的详细信息。")
        
        # 筛选器
        filter_col1, filter_col2 = st.columns(2)
        
        # 所有项目列表
        all_items = sorted(list(affected_items))
        with filter_col1:
            selected_items = st.multiselect("按项目筛选", all_items, default=all_items)
        with filter_col2:
            all_types = sorted(list(set(a["类型"] for a in anomalies)))
            selected_types = st.multiselect("按类型筛选", all_types, default=all_types)
        
        # 筛选
        filtered = [a for a in anomalies if a["项目"] in selected_items and a["类型"] in selected_types]
        
        # 严重程度排序
        type_order = {"🔴 严重异常": 0, "🟡 轻度异常": 1, "🔵 逻辑异常": 2, "⬆️ 历史最高值": 3, "⬇️ 历史最低值": 4}
        filtered.sort(key=lambda x: type_order.get(x.get("类型", ""), 9))
        
        if filtered:
            detail_rows = []
            for a in filtered:
                detail_rows.append({
                    "行号": a.get("行号", "-"),
                    "项目": a["项目"],
                    "类型": a["类型"],
                    "检测值": a.get("检测值", "-"),
                    "正常范围": a.get("正常范围", "-"),
                    "偏离程度": a.get("偏离程度", "-"),
                    "说明": a.get("说明", "-"),
                })
            detail_df = pd.DataFrame(detail_rows)
            st.dataframe(detail_df, use_container_width=True, hide_index=True, height=min(400, 35 * len(detail_rows) + 40))
            
            # 下载按钮
            csv = detail_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button("📥 导出异常报告 (CSV)", csv, "anomaly_report.csv", "text/csv")
        else:
            st.info("当前筛选条件下无异常记录")


# ==================== 页面7：模型管理 ====================
elif page == "📁 模型管理":
    st.header("📁 模型管理")
    
    saved = get_saved_models(st.session_state.current_user)

    if not saved:
        st.info("暂无已保存的模型，请先上传数据并训练")
    else:
        st.subheader(f"已保存的模型（{len(saved)}个）")
        
        for sm in saved:
            with st.container():
                col1, col2, col3 = st.columns([4, 1, 1])
                with col1:
                    group_info = f" | 分类：{sm['group']}" if sm['group'] != '全部' else ""
                    st.markdown(f"**{sm['display_name']}**{group_info} | 样本量：{sm['sample_count']} | 项目数：{sm['item_count']} | 相关对：{sm['pair_count']} | 训练：{sm['trained_at']}")
                with col2:
                    if st.button("加载", key=f"load_{sm['filename']}"):
                        model = load_model_file(sm["filepath"])
                        st.session_state.models[sm["name"]] = model
                        st.success(f"已加载模型：{sm['display_name']}")
                with col3:
                    if st.button("删除", key=f"del_{sm['filename']}"):
                        os.remove(sm["filepath"])
                        if sm["name"] in st.session_state.models:
                            del st.session_state.models[sm["name"]]
                        st.success(f"已删除模型：{sm['display_name']}")
                        st.rerun()
        
        # 查看已加载的模型详情
        if st.session_state.models:
            st.markdown("---")
            st.subheader("已加载的模型")
            for name, model in st.session_state.models.items():
                display_name = get_model_display_name(name, model, st.session_state.current_user)
                with st.expander(f"📊 {display_name}"):
                    stats_data = []
                    for col, info in model.get("item_stats", {}).items():
                        stats_data.append({
                            "项目": col,
                            "样本量": info['count'],
                            "均值": info['mean'],
                            "标准差": info['std'],
                            "正常范围": f"{info['lower_fence']} ~ {info['upper_fence']}"
                        })
                    if stats_data:
                        st.dataframe(pd.DataFrame(stats_data), use_container_width=True, hide_index=True)
                    
                    if model.get("strong_pairs"):
                        st.markdown("**强相关对：**")
                        for pair in model["strong_pairs"]:
                            icon = "📈" if pair["direction"] == "正相关" else "📉"
                            st.markdown(f"{icon} {pair['col_a']} ↔ {pair['col_b']}：{pair['direction']}，r={pair['r']}")
        
        # 重置模型
        st.markdown("---")
        st.subheader("⚠️ 重置操作")
        st.markdown("如果设备更换、煤种变化或数据出现系统性偏差，可以重置模型从头开始。")
        
        col_reset1, col_reset2 = st.columns(2)
        with col_reset1:
            if st.button("🔄 重置所有模型和数据", type="secondary"):
                st.session_state.models = {}
                st.session_state.df_raw = None
                st.session_state.numeric_cols = []
                st.session_state.category_cols = []
                st.session_state.decimal_settings = {}
                # 删除当前用户所有模型文件
                model_dir = get_model_dir(st.session_state.current_user)
                if os.path.exists(model_dir):
                    for f in os.listdir(model_dir):
                        if f.endswith(".json"):
                            os.remove(os.path.join(model_dir, f))
                st.success("✅ 已重置所有模型和数据")
                st.rerun()
        
        with col_reset2:
            if st.button("🗑️ 仅重置历史数据（保留模型设置）", type="secondary"):
                st.session_state.df_raw = None
                st.session_state.numeric_cols = []
                st.session_state.category_cols = []
                st.success("✅ 已重置历史数据，模型设置保留")
                st.rerun()

# ==================== 页面9：功能反馈 ====================
elif page == "💬 功能反馈":
    st.header("💬 功能反馈")
    st.markdown("您的需求驱动产品进化。告诉我们您需要什么功能，我们会优先开发。")
    
    # 反馈文件
    FEEDBACK_DIR = "feedback"
    os.makedirs(FEEDBACK_DIR, exist_ok=True)
    
    feedback_type = st.selectbox("反馈类型", [
        "💡 功能建议 — 我需要XXX功能",
        "🐛 问题反馈 — XXX用不了/有bug",
        "🤔 使用咨询 — 不知道怎么用XXX",
        "🌟 体验好评 — XXX做得不错"
    ])
    
    feedback_title = st.text_input("标题（简短描述）", placeholder="比如：希望增加SPC控制图功能")
    feedback_detail = st.text_area("详细描述", placeholder="请详细描述您的需求或问题，越具体越好...\n\n比如：\n- 你在什么场景下需要这个功能？\n- 期望的操作流程是什么？\n- 现在的工具有什么不方便的地方？", height=200)
    feedback_contact = st.text_input("联系方式（选填）", placeholder="方便我们跟进反馈")
    
    if st.button("📤 提交反馈", type="primary"):
        if not feedback_title:
            st.warning("请填写标题")
        elif not feedback_detail:
            st.warning("请填写详细描述")
        else:
            feedback = {
                "type": feedback_type,
                "title": feedback_title,
                "detail": feedback_detail,
                "contact": feedback_contact,
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            # 保存到文件
            filename = f"feedback_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            filepath = os.path.join(FEEDBACK_DIR, filename)
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(feedback, f, ensure_ascii=False, indent=2)
            
            st.success("✅ 感谢您的反馈！我们会认真对待每一条建议。")
            st.balloons()
    
    # 已提交的反馈
    st.markdown("---")
    st.subheader("📋 已提交的反馈")
    
    feedback_files = [f for f in os.listdir(FEEDBACK_DIR) if f.startswith("feedback_") and f.endswith(".json")] if os.path.exists(FEEDBACK_DIR) else []
    if feedback_files:
        for ff in sorted(feedback_files, reverse=True)[:20]:
            with open(os.path.join(FEEDBACK_DIR, ff), "r", encoding="utf-8") as f:
                fb = json.load(f)
            with st.expander(f"{fb['type']} — {fb['title']}（{fb['time']}）"):
                st.markdown(f"**详细描述：**\n{fb['detail']}")
                if fb.get("contact"):
                    st.caption(f"联系方式：{fb['contact']}")
    else:
        st.info("暂无反馈记录")
    
    # 常见需求投票
    st.markdown("---")
    st.subheader("🗳️ 功能需求投票")
    st.markdown("以下功能正在规划中，您最想要哪些？")
    
    planned_features = [
        ("📷 图片OCR上传", "拍照上传检测报告，自动识别数据"),
        ("📱 微信/邮件异常推送", "检测到异常时自动推送通知"),
        ("📊 SPC控制图", "行业标准统计过程控制图（X-bar、R图等）"),
        ("📈 行业基准对比", "与行业平均水平对比，了解自身位置"),
        ("📄 检测报告导出", "一键导出PDF/Word格式的检测报告"),
        ("🔐 账号体系", "多用户独立工作区，数据隔离"),
        ("🔄 连续异常预警", "连续N天异常自动升级预警"),
        ("⚖️ 项目综合健康评分", "给每次检测打0-100分，一眼判断整体质量"),
        ("📉 批次间对比", "本次采样与上次对比，发现变化最大的项目")
    ]
    
    voted = st.session_state.get("feature_votes", [])
    
    for fname, fdesc in planned_features:
        col1, col2 = st.columns([4, 1])
        with col1:
            st.markdown(f"**{fname}**")
            st.caption(fdesc)
        with col2:
            if fname in voted:
                st.markdown("✅ 已投票")
            else:
                if st.button("👍 投票", key=f"vote_{fname}"):
                    voted.append(fname)
                    st.session_state.feature_votes = voted
                    # 保存投票
                    vote_file = os.path.join(FEEDBACK_DIR, "votes.json")
                    votes = {}
                    if os.path.exists(vote_file):
                        with open(vote_file, "r", encoding="utf-8") as f:
                            votes = json.load(f)
                    votes[fname] = votes.get(fname, 0) + 1
                    with open(vote_file, "w", encoding="utf-8") as f:
                        json.dump(votes, f, ensure_ascii=False, indent=2)
                    st.success(f"已投票：{fname}")
                    st.rerun()
    
    # 投票结果
    vote_file = os.path.join(FEEDBACK_DIR, "votes.json")
    if os.path.exists(vote_file):
        with open(vote_file, "r", encoding="utf-8") as f:
            votes = json.load(f)
        if votes:
            st.markdown("---")
            st.subheader("📊 当前投票排行")
            sorted_votes = sorted(votes.items(), key=lambda x: x[1], reverse=True)
            for fname, count in sorted_votes:
                st.markdown(f"- **{fname}**：{count} 票")

# ==================== 页面10：用户管理（管理员） ====================
elif page == "⚙️ 用户管理":
    if not is_admin:
        st.error("⛔ 仅管理员可访问此页面")
    else:
        st.header("⚙️ 用户管理")
        
        # ===== 动态口令设置 =====
        st.subheader("🔐 动态口令（TOTP）设置")
        my_data = current_user_data
        
        if my_data.get("totp_enabled", False):
            st.success("✅ 动态口令验证已开启 — 登录时需输入验证器App的6位数字")
            if st.button("❌ 关闭动态口令", type="secondary"):
                my_data["totp_enabled"] = False
                my_data["totp_secret"] = ""
                save_user(my_data)
                st.success("已关闭动态口令验证")
                st.rerun()
        else:
            st.warning("⚠️ 动态口令未开启 — 建议管理员开启，提高账号安全性")
            st.markdown("**开启步骤：**")
            st.markdown("1. 手机安装任意TOTP验证器App（如 Microsoft Authenticator、身份验证器等）")
            st.markdown("2. 点击下方按钮生成密钥")
            st.markdown("3. 用App扫描二维码或手动输入密钥")
            st.markdown("4. 输入App显示的6位数字完成绑定")
            
            if st.button("🔄 生成动态口令密钥", type="primary"):
                secret = generate_totp_secret()
                st.session_state["temp_totp_secret"] = secret
            
            temp_secret = st.session_state.get("temp_totp_secret")
            if temp_secret:
                qr_b64 = get_totp_qr_base64(my_data["username"], temp_secret)
                if qr_b64:
                    st.markdown(f"""
                    <div style="text-align:center; padding:16px; background:white; border-radius:12px; margin:16px 0;">
                        <img src="data:image/png;base64,{qr_b64}" width="200"/>
                        <div style="color:#333; font-size:12px; margin-top:8px;">用验证器App扫描此二维码</div>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.info("qrcode库未安装，请手动输入以下密钥：")
                
                st.code(temp_secret, language=None)
                st.caption("如果无法扫码，在App中手动输入上方密钥")
                
                bind_code = st.text_input("输入验证器App显示的6位数字", placeholder="000000", key="bind_totp_code")
                if st.button("✅ 确认绑定", type="primary"):
                    if not bind_code:
                        st.warning("请输入6位数字验证码")
                    elif verify_totp(temp_secret, bind_code):
                        my_data["totp_enabled"] = True
                        my_data["totp_secret"] = temp_secret
                        save_user(my_data)
                        st.session_state["temp_totp_secret"] = None
                        st.success("✅ 动态口令绑定成功！下次登录需输入验证码")
                        st.rerun()
                    else:
                        st.error("验证码错误，请确认App中的数字")
        
        st.markdown("---")
        
        # ===== 用户列表 =====
        st.subheader("👥 用户列表")
        all_users = get_all_users()
        
        if not all_users:
            st.info("暂无注册用户")
        else:
            admin_count = len([u for u in all_users if u.get("role") == "admin"])
            user_count = len([u for u in all_users if u.get("role") == "user"])
            banned_count = len([u for u in all_users if u.get("banned", False)])
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("总用户", f"{len(all_users)}")
            with col2:
                st.metric("管理员", f"{admin_count}")
            with col3:
                st.metric("普通用户", f"{user_count}")
            with col4:
                st.metric("已封禁", f"{banned_count}")
            
            st.markdown("---")
            
            for u in all_users:
                role_icon = "👑" if u.get("role") == "admin" else "👤"
                status = "🚫 已封禁" if u.get("banned", False) else "✅ 正常"
                totp_status = "🔐 已开启" if u.get("totp_enabled", False) else "—"
                is_me = u["username"] == st.session_state.current_user
                
                with st.expander(f"{role_icon} {u['username']} {'（当前账号）' if is_me else ''} — {status}"):
                    col_info, col_action = st.columns([3, 1])
                    with col_info:
                        st.markdown(f"""
                        - **角色：** {u.get('role', 'user')}
                        - **邮箱：** {u.get('email', '—')}
                        - **注册时间：** {u.get('created_at', '—')}
                        - **最后登录：** {u.get('last_login', '—')}
                        - **动态口令：** {totp_status}
                        - **失败尝试：** {u.get('failed_attempts', 0)}次
                        """)
                        
                        if u.get("locked_until", 0) > time.time():
                            remaining = int(u["locked_until"] - time.time())
                            st.warning(f"🔒 账号锁定中，剩余 {remaining // 60 + 1} 分钟")
                    
                    with col_action:
                        if not is_me:
                            if u.get("banned", False):
                                if st.button("✅ 解封", key=f"unban_{u['username']}"):
                                    u["banned"] = False
                                    u["failed_attempts"] = 0
                                    u["locked_until"] = 0
                                    save_user(u)
                                    st.success(f"已解封 {u['username']}")
                                    st.rerun()
                            else:
                                if st.button("🚫 封禁", key=f"ban_{u['username']}"):
                                    u["banned"] = True
                                    save_user(u)
                                    st.success(f"已封禁 {u['username']}")
                                    st.rerun()
                            
                            if st.button("🔑 重置密码", key=f"reset_{u['username']}"):
                                new_pass = f"szzk{datetime.now().strftime('%m%d')}"
                                salt = u.get("salt", generate_salt())
                                u["salt"] = salt
                                u["password_hash"] = hash_password(new_pass, salt)
                                u["failed_attempts"] = 0
                                u["locked_until"] = 0
                                save_user(u)
                                st.success(f"已重置 {u['username']} 的密码为：`{new_pass}`，请通知用户尽快修改")
                        else:
                            st.caption("当前账号")
        
        st.markdown("---")
        
        # ===== 修改密码 =====
        st.subheader("🔑 修改我的密码")
        old_pass = st.text_input("当前密码", type="password", key="admin_old_pass")
        new_pass1 = st.text_input("新密码", type="password", placeholder="6个字符以上", key="admin_new_pass1")
        new_pass2 = st.text_input("确认新密码", type="password", key="admin_new_pass2")
        if st.button("💾 修改密码", type="primary"):
            if not old_pass or not new_pass1:
                st.warning("请填写密码")
            elif new_pass1 != new_pass2:
                st.warning("两次密码不一致")
            elif len(new_pass1) < 6:
                st.warning("新密码至少6个字符")
            else:
                salt = my_data.get("salt", "")
                if my_data["password_hash"] != hash_password(old_pass, salt):
                    st.error("当前密码错误")
                else:
                    new_salt = generate_salt()
                    my_data["salt"] = new_salt
                    my_data["password_hash"] = hash_password(new_pass1, new_salt)
                    save_user(my_data)
                    st.success("✅ 密码修改成功")