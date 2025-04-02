import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import datetime
import calendar
import requests
import io
import warnings
from typing import Dict, List, Tuple, Union, Optional
from io import BytesIO
import base64

warnings.filterwarnings('ignore')

# 设置页面配置
st.set_page_config(
    page_title="物料投放分析动态仪表盘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 全局样式
st.markdown("""
<style>
    .main-header {color:#1E88E5; font-size:28px; font-weight:bold; text-align:center;}
    .sub-header {color:#424242; font-size:20px; font-weight:bold; margin-top:12px;}
    .metric-card {background-color:#f5f5f5; border-radius:5px; padding:10px; box-shadow:1px 1px 3px #cccccc;}
    .tooltip {position:relative; display:inline-block; cursor:help;}
    .tooltip .tooltiptext {visibility:hidden; width:300px; background-color:#555; color:#fff; text-align:left; 
        border-radius:6px; padding:10px; position:absolute; z-index:1; bottom:125%; left:50%; margin-left:-150px; 
        opacity:0; transition:opacity 0.3s;}
    .tooltip:hover .tooltiptext {visibility:visible; opacity:1;}
    .highlight-box {border-left:3px solid #1E88E5; padding-left:10px; margin:10px 0px;}
    .success-metric {color:#4CAF50;}
    .warning-metric {color:#FF9800;}
    .danger-metric {color:#F44336;}
    .info-box {background-color:#E3F2FD; padding:10px; border-radius:5px; margin:10px 0px;}
    .spacer {height:20px;}
</style>
""", unsafe_allow_html=True)

# 业务指标定义字典
BUSINESS_DEFINITIONS = {
    "物料总成本": "指所有物料的成本总和，计算方式为：物料数量 × 物料单价。用于衡量物料投入的总金额。",
    "销售总额": "指所有产品的销售收入总和，计算方式为：产品数量 × 产品单价。用于衡量销售业绩。",
    "投资回报率(ROI)": "投资回报率，计算方式为：销售总额 ÷ 物料总成本。ROI > 1表示物料投入产生了正回报，ROI < 1表示投入未获得有效回报。",
    "物料销售比率": "物料总成本占销售总额的百分比，计算方式为：物料总成本 ÷ 销售总额 × 100%。该比率越低，表示物料使用效率越高。",
    "高效物料投放经销商": "ROI值大于行业平均水平(通常为2.0)的经销商，这些经销商能够高效地利用物料创造销售。",
    "待优化物料投放经销商": "ROI值低于1.0的经销商，这些经销商的物料使用效率有待提高，物料投入未能产生等价销售回报。",
    "物料使用效率": "衡量单位物料投入所产生的销售额，计算方式为：销售额 ÷ 物料数量。效率越高，表示物料利用度越好。",
    "物料覆盖率": "使用某种物料的经销商数量占总经销商数量的百分比，用于评估物料普及度。",
    "销售转化周期": "从物料投放到产生销售的平均时间间隔，用于评估物料效果显现的速度。",
    "客户价值分层": "基于物料ROI和销售额将客户分为高价值、成长型、稳定型和低效型四类，用于差异化管理。"
}

# GitHub相关配置
github_owner = "你的GitHub用户名"  # 替换为您的GitHub用户名
github_repo = "物料投放分析"  # 替换为您的GitHub仓库名
github_branch = "main"  # GitHub分支，通常是main或master

# GitHub文件路径
github_files = {
    "material_data": "2025物料源数据.xlsx",
    "sales_data": "25物料源销售数据.xlsx",
    "material_price": "物料单价.xlsx"
}

# 本地文件路径（备用）
local_files = {
    "material_data": r"C:\Users\何晴雅\Desktop\2025物料源数据.xlsx",
    "sales_data": r"C:\Users\何晴雅\Desktop\25物料源销售数据.xlsx",
    "material_price": r"C:\Users\何晴雅\Desktop\物料单价.xlsx"
}


# 格式化金额函数 - 确保所有金额都保留两位小数
def format_currency(value: float) -> str:
    """格式化金额为带两位小数的字符串"""
    return f"{value:.2f}元"


# 创建下载按钮
def create_download_link(df, filename):
    """为DataFrame创建下载链接"""
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}.csv">下载报表</a>'
    return href


# 创建带解释的工具提示
def create_tooltip(text, explanation):
    """创建带解释的工具提示"""
    return f"""
    <div class="tooltip">{text}
        <span class="tooltiptext">{explanation}</span>
    </div>
    """


# 从GitHub加载文件
def load_from_github(file_path, github_owner, github_repo, github_branch):
    """从GitHub加载Excel文件"""
    raw_url = f"https://raw.githubusercontent.com/{github_owner}/{github_repo}/{github_branch}/{file_path}"
    try:
        response = requests.get(raw_url)
        response.raise_for_status()  # 确保请求成功
        return pd.read_excel(io.BytesIO(response.content))
    except Exception as e:
        st.error(f"从GitHub加载文件失败: {e}")
        return None


# 数据加载与处理
@st.cache_data
def load_data(use_github=True):
    """加载和处理数据"""
    # 尝试从GitHub加载数据，如果失败则使用本地文件
    if use_github:
        try:
            # 加载物料源数据
            material_data = load_from_github(github_files["material_data"], github_owner, github_repo, github_branch)

            # 加载销售数据
            sales_data = load_from_github(github_files["sales_data"], github_owner, github_repo, github_branch)

            # 加载物料单价数据
            material_price = load_from_github(github_files["material_price"], github_owner, github_repo, github_branch)

            # 检查是否所有数据都加载成功
            if material_data is None or sales_data is None or material_price is None:
                raise Exception("部分数据加载失败，将使用本地文件")

        except Exception as e:
            st.warning(f"从GitHub加载数据失败，将使用本地文件: {e}")
            use_github = False

    # 如果GitHub加载失败或者选择使用本地文件，则加载本地文件
    if not use_github:
        # 加载物料源数据
        material_data = pd.read_excel(local_files["material_data"])

        # 加载销售数据
        sales_data = pd.read_excel(local_files["sales_data"])

        # 加载物料单价数据
        material_price = pd.read_excel(local_files["material_price"])

    # 数据处理部分（保持不变，但添加了更多衍生指标）
    # 处理日期格式
    material_data['发运月份'] = pd.to_datetime(material_data['发运月份'])
    sales_data['发运月份'] = pd.to_datetime(sales_data['发运月份'])

    # 创建月份和年份列
    for df in [material_data, sales_data]:
        df['月份'] = df['发运月份'].dt.month
        df['年份'] = df['发运月份'].dt.year
        df['月份名'] = df['发运月份'].dt.strftime('%Y-%m')
        df['季度'] = df['发运月份'].dt.quarter  # 新增季度列
        df['月度名称'] = df['发运月份'].dt.strftime('%m月')  # 新增月度名称，便于分析

    # 计算物料成本
    material_data = pd.merge(material_data, material_price[['物料代码', '单价（元）', '物料类别']],
                             left_on='产品代码', right_on='物料代码', how='left')

    # 填充缺失的物料单价为平均值
    mean_price = material_price['单价（元）'].mean()
    material_data['单价（元）'].fillna(mean_price, inplace=True)

    # 计算物料总成本
    material_data['物料成本'] = material_data['求和项:数量（箱）'] * material_data['单价（元）']

    # 计算销售总金额
    sales_data['销售金额'] = sales_data['求和项:数量（箱）'] * sales_data['求和项:单价（箱）']

    # 新增：计算单位物料使用效率
    material_data['单位使用效率'] = material_data['求和项:数量（箱）'] / material_data['物料成本'].replace(0, np.nan)
    material_data['单位使用效率'].fillna(0, inplace=True)

    # 按经销商、月份计算物料成本总和
    material_cost_by_distributor = material_data.groupby(['客户代码', '经销商名称', '月份名'])[
        '物料成本'].sum().reset_index()
    material_cost_by_distributor.rename(columns={'物料成本': '物料总成本'}, inplace=True)

    # 按经销商、月份计算销售总额
    sales_by_distributor = sales_data.groupby(['客户代码', '经销商名称', '月份名'])['销售金额'].sum().reset_index()
    sales_by_distributor.rename(columns={'销售金额': '销售总额'}, inplace=True)

    # 合并物料成本和销售数据
    distributor_data = pd.merge(material_cost_by_distributor, sales_by_distributor,
                                on=['客户代码', '经销商名称', '月份名'], how='outer').fillna(0)

    # 计算ROI
    distributor_data['ROI'] = np.where(distributor_data['物料总成本'] > 0,
                                       distributor_data['销售总额'] / distributor_data['物料总成本'], 0)

    # 计算物料销售比率
    distributor_data['物料销售比率'] = (distributor_data['物料总成本'] / distributor_data['销售总额'].replace(0,
                                                                                                              np.nan)) * 100
    distributor_data['物料销售比率'].fillna(0, inplace=True)

    # 新增：经销商价值分层
    def value_segment(row):
        if row['ROI'] >= 2.0 and row['销售总额'] > distributor_data['销售总额'].quantile(0.75):
            return '高价值客户'
        elif row['ROI'] >= 1.0 and row['销售总额'] > distributor_data['销售总额'].median():
            return '成长型客户'
        elif row['ROI'] >= 1.0:
            return '稳定型客户'
        else:
            return '低效型客户'

    distributor_data['客户价值分层'] = distributor_data.apply(value_segment, axis=1)

    # 新增：物料使用多样性（经销商使用的物料种类数）
    material_diversity = material_data.groupby(['客户代码', '月份名'])['产品代码'].nunique().reset_index()
    material_diversity.rename(columns={'产品代码': '物料多样性'}, inplace=True)

    # 合并物料多样性到经销商数据
    distributor_data = pd.merge(distributor_data, material_diversity,
                                on=['客户代码', '月份名'], how='left')
    distributor_data['物料多样性'].fillna(0, inplace=True)

    return material_data, sales_data, material_price, distributor_data


# 主应用
def main():
    """主应用函数"""
    st.markdown('<h1 class="main-header">物料投放分析动态仪表盘</h1>', unsafe_allow_html=True)
    st.markdown(
        '<p style="text-align:center; font-size:18px;">协助销售人员更有效地利用物料资源，优化投放策略，提升销售业绩</p>',
        unsafe_allow_html=True)

    # 加载数据
    material_data, sales_data, material_price, distributor_data = load_data(use_github=False)  # 默认使用本地文件

    # 侧边栏筛选条件
    st.sidebar.markdown('## 筛选条件')

    # 区域列表
    regions = sorted(material_data['所属区域'].unique())
    selected_regions = st.sidebar.multiselect("选择区域:", regions, default=regions)

    # 省份列表
    provinces = sorted(material_data['省份'].unique())
    selected_provinces = st.sidebar.multiselect("选择省份:", provinces, default=provinces)

    # 月份列表
    months = sorted(material_data['月份名'].unique())
    selected_month = st.sidebar.selectbox("选择月份:", months)

    # 物料类别列表
    material_categories = sorted(material_price['物料类别'].unique())
    selected_categories = st.sidebar.multiselect("选择物料类别:", material_categories, default=material_categories)

    # 新增：价值分层筛选
    value_segments = distributor_data['客户价值分层'].unique()
    selected_segments = st.sidebar.multiselect("选择客户价值分层:", value_segments, default=value_segments)

    # 更新按钮
    update_button = st.sidebar.button("更新仪表盘")

    # 侧边栏指标说明
    st.sidebar.markdown('## 业务指标说明')

    for term, definition in BUSINESS_DEFINITIONS.items():
        with st.sidebar.expander(term):
            st.write(definition)

    # 筛选数据
    if update_button or True:  # 默认自动更新
        # 筛选物料数据
        filtered_material = material_data[
            (material_data['所属区域'].isin(selected_regions)) &
            (material_data['省份'].isin(selected_provinces)) &
            (material_data['月份名'] == selected_month)
            ]

        # 合并物料类别信息并筛选
        filtered_material = filtered_material[filtered_material['物料类别'].isin(selected_categories)]

        # 筛选销售数据
        filtered_sales = sales_data[
            (sales_data['所属区域'].isin(selected_regions)) &
            (sales_data['省份'].isin(selected_provinces)) &
            (sales_data['月份名'] == selected_month)
            ]

        # 筛选经销商数据
        filtered_distributor = distributor_data[
            distributor_data['月份名'] == selected_month
            ]

        # 进一步筛选经销商（按价值分层）
        filtered_distributor = filtered_distributor[filtered_distributor['客户价值分层'].isin(selected_segments)]

        # 通过客户代码筛选经销商数据，确保只保留符合区域和省份条件的经销商
        valid_distributors = filtered_sales['客户代码'].unique()
        filtered_distributor = filtered_distributor[filtered_distributor['客户代码'].isin(valid_distributors)]

        # 创建标签页
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "业绩概览",
            "物料与销售分析",
            "经销商分析",
            "ROI分析",
            "客户价值分层",
            "季节性分析",
            "优化建议"
        ])

        # ======= 业绩概览标签页 =======
        with tab1:
            st.markdown(
                '<div class="info-box">本页面显示关键业绩指标汇总和主要分析图表，帮助您快速了解物料投放效果和整体销售业绩表现。</div>',
                unsafe_allow_html=True)

            # 创建汇总统计信息
            total_material_cost = filtered_material['物料成本'].sum()
            total_sales = filtered_sales['销售金额'].sum()
            roi = total_sales / total_material_cost if total_material_cost > 0 else 0
            total_distributors = filtered_sales['经销商名称'].nunique()
            material_sales_ratio = (total_material_cost / total_sales * 100) if total_sales > 0 else 0

            # 平均每个经销商物料成本和销售额
            avg_material_cost = total_material_cost / total_distributors if total_distributors > 0 else 0
            avg_sales = total_sales / total_distributors if total_distributors > 0 else 0

            # 使用Streamlit的度量标准组件
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("物料总成本", format_currency(total_material_cost))
                st.caption(f"平均每个经销商: {format_currency(avg_material_cost)}")
            with col2:
                st.metric("销售总额", format_currency(total_sales))
                st.caption(f"平均每个经销商: {format_currency(avg_sales)}")
            with col3:
                # 添加ROI颜色指示
                roi_color = "success-metric" if roi >= 2.0 else "warning-metric" if roi >= 1.0 else "danger-metric"
                st.markdown(f'<div class="metric-card"><h3 class="{roi_color}">ROI: {roi:.2f}</h3></div>',
                            unsafe_allow_html=True)
                st.caption("销售总额 ÷ 物料总成本")
            with col4:
                # 添加物料销售比率颜色指示
                ratio_color = "success-metric" if material_sales_ratio <= 30 else "warning-metric" if material_sales_ratio <= 50 else "danger-metric"
                st.markdown(
                    f'<div class="metric-card"><h3 class="{ratio_color}">物料销售比率: {material_sales_ratio:.2f}%</h3></div>',
                    unsafe_allow_html=True)
                st.caption("物料总成本 ÷ 销售总额 × 100%")

            st.markdown('<div class="spacer"></div>', unsafe_allow_html=True)

            # 新增：整体业绩仪表盘
            st.markdown('<h3 class="sub-header">整体业绩仪表盘</h3>', unsafe_allow_html=True)

            # 创建仪表盘指标
            gauge_cols = st.columns(3)

            with gauge_cols[0]:
                roi_gauge = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=roi,
                    title={'text': "投资回报率(ROI)"},
                    gauge={
                        'axis': {'range': [None, 5], 'tickwidth': 1, 'tickcolor': "darkblue"},
                        'bar': {'color': "darkblue"},
                        'bgcolor': "white",
                        'borderwidth': 2,
                        'bordercolor': "gray",
                        'steps': [
                            {'range': [0, 1], 'color': '#FF9E9E'},
                            {'range': [1, 2], 'color': '#FFEC9E'},
                            {'range': [2, 5], 'color': '#9EFF9E'}
                        ],
                        'threshold': {
                            'line': {'color': "red", 'width': 4},
                            'thickness': 0.75,
                            'value': 1
                        }
                    }
                ))
                roi_gauge.update_layout(height=300, margin=dict(l=20, r=20, t=50, b=20))
                st.plotly_chart(roi_gauge, use_container_width=True)

            with gauge_cols[1]:
                high_value_percent = (filtered_distributor['客户价值分层'] == '高价值客户').mean() * 100
                customer_gauge = go.Figure(go.Indicator(
                    mode="gauge+number+delta",
                    value=high_value_percent,
                    title={'text': "高价值客户占比(%)"},
                    delta={'reference': 25, 'increasing': {'color': "green"}},
                    gauge={
                        'axis': {'range': [None, 100], 'tickwidth': 1},
                        'bar': {'color': "darkblue"},
                        'steps': [
                            {'range': [0, 20], 'color': '#FFB6C1'},
                            {'range': [20, 40], 'color': '#FFFACD'},
                            {'range': [40, 100], 'color': '#90EE90'}
                        ],
                        'threshold': {
                            'line': {'color': "red", 'width': 4},
                            'thickness': 0.75,
                            'value': 25
                        }
                    }
                ))
                customer_gauge.update_layout(height=300, margin=dict(l=20, r=20, t=50, b=20))
                st.plotly_chart(customer_gauge, use_container_width=True)

            with gauge_cols[2]:
                material_gauge = go.Figure(go.Indicator(
                    mode="gauge+number+delta",
                    value=material_sales_ratio,
                    title={'text': "物料销售比率(%)"},
                    delta={'reference': 50, 'decreasing': {'color': "green"}},
                    gauge={
                        'axis': {'range': [None, 100], 'tickwidth': 1},
                        'bar': {'color': "darkblue"},
                        'steps': [
                            {'range': [0, 30], 'color': '#90EE90'},
                            {'range': [30, 50], 'color': '#FFFACD'},
                            {'range': [50, 100], 'color': '#FFB6C1'}
                        ],
                        'threshold': {
                            'line': {'color': "red", 'width': 4},
                            'thickness': 0.75,
                            'value': 50
                        }
                    }
                ))
                material_gauge.update_layout(height=300, margin=dict(l=20, r=20, t=50, b=20))
                st.plotly_chart(material_gauge, use_container_width=True)

            st.markdown(
                '<div class="highlight-box">以上仪表盘显示了关键业绩指标的当前状态。ROI目标值为2.0以上，高价值客户占比目标为25%以上，物料销售比率目标为30%以下。</div>',
                unsafe_allow_html=True)

            # ROI图表和热门物料图表
            col1, col2 = st.columns(2)

            with col1:
                # ROI图表
                roi_by_region = filtered_distributor.groupby('客户代码').agg({
                    '物料总成本': 'sum',
                    '销售总额': 'sum'
                }).reset_index()

                roi_by_region['ROI'] = roi_by_region['销售总额'] / roi_by_region['物料总成本'].replace(0, np.nan)
                roi_by_region['ROI'].fillna(0, inplace=True)

                # 合并经销商名称
                distributor_names = filtered_distributor[['客户代码', '经销商名称']].drop_duplicates()
                roi_by_region = pd.merge(roi_by_region, distributor_names, on='客户代码')

                # 按ROI排序并取前10名
                top_roi = roi_by_region.sort_values('ROI', ascending=False).head(10)

                roi_chart = px.bar(
                    top_roi,
                    x='经销商名称',
                    y='ROI',
                    title='ROI最高的经销商Top 10',
                    color='ROI',
                    color_continuous_scale='Viridis',
                    text='ROI',
                    hover_data={
                        '物料总成本': ':.2f',
                        '销售总额': ':.2f',
                        'ROI': ':.2f'
                    }
                )
                roi_chart.update_traces(texttemplate='%{text:.2f}', textposition='outside')
                roi_chart.update_layout(
                    xaxis_tickangle=-45,
                    xaxis_title="经销商名称",
                    yaxis_title="投资回报率(ROI)"
                )

                st.plotly_chart(roi_chart, use_container_width=True)
                st.markdown(
                    '<div class="highlight-box">上图显示ROI最高的经销商，这些经销商能够高效利用物料创造销售，可以研究其物料使用策略用于指导其他经销商。</div>',
                    unsafe_allow_html=True)

            with col2:
                # 热门物料图表 - 添加物料类别颜色分组
                top_materials = filtered_material.groupby(['产品名称', '物料类别'])[
                    '求和项:数量（箱）'].sum().reset_index()
                top_materials = top_materials.sort_values('求和项:数量（箱）', ascending=False).head(10)

                top_materials_chart = px.bar(
                    top_materials,
                    x='产品名称',
                    y='求和项:数量（箱）',
                    title='最热门物料Top 10 (按数量)',
                    color='物料类别',
                    text='求和项:数量（箱）'
                )
                top_materials_chart.update_traces(texttemplate='%{text:.0f}', textposition='outside')
                top_materials_chart.update_layout(
                    xaxis_tickangle=-45,
                    xaxis_title="物料名称",
                    yaxis_title="数量（箱）"
                )

                st.plotly_chart(top_materials_chart, use_container_width=True)
                st.markdown(
                    '<div class="highlight-box">上图显示发放数量最多的物料及其类别，帮助了解哪些物料最受欢迎，是否有特定物料类别更受青睐。</div>',
                    unsafe_allow_html=True)

            # 区域比较图表
            st.markdown('<h3 class="sub-header">区域业绩对比</h3>', unsafe_allow_html=True)

            region_comparison = filtered_material.groupby('所属区域').agg({
                '物料成本': 'sum'
            }).reset_index()

            sales_by_region = filtered_sales.groupby('所属区域').agg({
                '销售金额': 'sum'
            }).reset_index()

            region_comparison = pd.merge(region_comparison, sales_by_region, on='所属区域', how='outer').fillna(0)
            region_comparison['ROI'] = region_comparison['销售金额'] / region_comparison['物料成本'].replace(0, np.nan)
            region_comparison['ROI'].fillna(0, inplace=True)
            region_comparison['物料销售比率'] = (
                        region_comparison['物料成本'] / region_comparison['销售金额'] * 100).replace(np.inf, 0).fillna(
                0)

            region_comparison_chart = make_subplots(specs=[[{"secondary_y": True}]])

            region_comparison_chart.add_trace(
                go.Bar(x=region_comparison['所属区域'],
                       y=region_comparison['物料成本'],
                       name='物料成本',
                       marker_color='rgba(58, 71, 80, 0.6)',
                       hovertemplate='区域: %{x}<br>物料成本: %{y:.2f}元<extra></extra>'),
                secondary_y=False
            )

            region_comparison_chart.add_trace(
                go.Bar(x=region_comparison['所属区域'],
                       y=region_comparison['销售金额'],
                       name='销售金额',
                       marker_color='rgba(246, 78, 139, 0.6)',
                       hovertemplate='区域: %{x}<br>销售金额: %{y:.2f}元<extra></extra>'),
                secondary_y=False
            )

            region_comparison_chart.add_trace(
                go.Scatter(x=region_comparison['所属区域'],
                           y=region_comparison['ROI'],
                           name='ROI',
                           mode='lines+markers+text',
                           line=dict(color='rgb(25, 118, 210)', width=3),
                           marker=dict(size=10),
                           text=region_comparison['ROI'].round(2),
                           textposition='top center',
                           hovertemplate='区域: %{x}<br>ROI: %{y:.2f}<extra></extra>'),
                secondary_y=True
            )

            region_comparison_chart.update_layout(
                title_text='区域比较: 物料成本、销售金额和ROI',
                barmode='group',
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                )
            )

            region_comparison_chart.update_yaxes(title_text='金额 (元)', secondary_y=False)
            region_comparison_chart.update_yaxes(title_text='ROI', secondary_y=True)

            st.plotly_chart(region_comparison_chart, use_container_width=True)

            # 添加区域表现评估表格
            region_comparison['表现评估'] = np.where(region_comparison['ROI'] >= 2.0, '优秀',
                                                     np.where(region_comparison['ROI'] >= 1.0, '良好', '需改进'))

            region_comparison['物料成本'] = region_comparison['物料成本'].round(2)
            region_comparison['销售金额'] = region_comparison['销售金额'].round(2)
            region_comparison['ROI'] = region_comparison['ROI'].round(2)
            region_comparison['物料销售比率'] = region_comparison['物料销售比率'].round(2)

            st.markdown('<h4>区域表现评估</h4>', unsafe_allow_html=True)
            st.dataframe(
                region_comparison[['所属区域', '物料成本', '销售金额', 'ROI', '物料销售比率', '表现评估']].rename(
                    columns={'物料销售比率': '物料销售比率(%)'}, inplace=False
                ),
                use_container_width=True,
                column_config={
                    "物料成本": st.column_config.NumberColumn("物料成本(元)", format="¥%.2f"),
                    "销售金额": st.column_config.NumberColumn("销售金额(元)", format="¥%.2f"),
                    "ROI": st.column_config.NumberColumn("ROI", format="%.2f"),
                    "物料销售比率(%)": st.column_config.NumberColumn("物料销售比率(%)", format="%.2f%%"),
                    "表现评估": st.column_config.TextColumn("表现评估", help="基于ROI的表现评估")
                }
            )

            st.markdown(
                f'<div class="highlight-box">此表显示了各区域的表现情况，ROI ≥ 2.0为优秀，1.0 ≤ ROI < 2.0为良好，ROI < 1.0为需改进。建议重点关注ROI较低的区域，分析原因并制定改进计划。</div>',
                unsafe_allow_html=True)

            # 提供下载区域业绩数据的链接
            st.markdown(create_download_link(region_comparison, "区域业绩数据"), unsafe_allow_html=True)

        # ======= 物料与销售分析标签页 =======
        with tab2:
            st.markdown(
                '<div class="info-box">本页面分析物料投入和销售产出的关系，帮助您了解物料投入与销售业绩的相关性并识别最有效的物料使用模式。</div>',
                unsafe_allow_html=True)

            # 物料-销售关系图
            st.markdown('<h3 class="sub-header">物料成本与销售金额关系</h3>', unsafe_allow_html=True)

            material_sales_relation = filtered_distributor.copy()

            # 添加颜色区分客户价值分层
            material_sales_chart = px.scatter(
                material_sales_relation,
                x='物料总成本',
                y='销售总额',
                size='ROI',
                color='客户价值分层',
                hover_name='经销商名称',
                log_x=True,
                log_y=True,
                title='物料成本与销售金额关系散点图',
                size_max=40,
                hover_data={
                    '物料总成本': ':.2f',
                    '销售总额': ':.2f',
                    'ROI': ':.2f',
                    '物料多样性': True
                }
            )

            material_sales_chart.update_layout(
                xaxis_title="物料总成本 (元，对数刻度)",
                yaxis_title="销售总额 (元，对数刻度)"
            )

            # 添加参考线 - ROI=1
            material_sales_chart.add_trace(
                go.Scatter(
                    x=[material_sales_relation['物料总成本'].min(), material_sales_relation['物料总成本'].max()],
                    y=[material_sales_relation['物料总成本'].min(), material_sales_relation['物料总成本'].max()],
                    mode='lines',
                    line=dict(color='red', dash='dash'),
                    name='ROI=1 参考线',
                    showlegend=True
                )
            )

            # 添加参考线 - ROI=2
            material_sales_chart.add_trace(
                go.Scatter(
                    x=[material_sales_relation['物料总成本'].min(), material_sales_relation['物料总成本'].max()],
                    y=[material_sales_relation['物料总成本'].min() * 2,
                       material_sales_relation['物料总成本'].max() * 2],
                    mode='lines',
                    line=dict(color='green', dash='dash'),
                    name='ROI=2 参考线',
                    showlegend=True
                )
            )

            st.plotly_chart(material_sales_chart, use_container_width=True)
            st.markdown(
                '<div class="highlight-box">此散点图展示物料成本与销售金额的关系，点的大小代表ROI，颜色代表客户价值分层。理想情况下，经销商应位于红线(ROI=1)以上，最好在绿线(ROI=2)以上，表示物料投入获得了较好的回报。</div>',
                unsafe_allow_html=True)

            # 物料使用模式分析
            st.markdown('<h3 class="sub-header">物料使用模式分析</h3>', unsafe_allow_html=True)

            # 按物料类别分析ROI
            material_category_roi = filtered_material.groupby('物料类别').agg({
                '物料成本': 'sum',
                '求和项:数量（箱）': 'sum'
            }).reset_index()

            # 匹配销售数据
            material_category_sales = filtered_sales.copy()

            # 为物料数据创建匹配代码，通过产品代码前缀匹配
            material_data_for_match = material_data[['产品代码', '物料类别']].drop_duplicates()

            # 提取产品代码前缀(假设前6个字符是匹配码)
            material_data_for_match['匹配代码'] = material_data_for_match['产品代码'].str[:6]
            material_category_sales['匹配代码'] = material_category_sales['产品代码'].str[:6]

            # 合并物料类别到销售数据
            material_category_sales = pd.merge(
                material_category_sales,
                material_data_for_match[['匹配代码', '物料类别']],
                on='匹配代码',
                how='left'
            )

            # 按物料类别汇总销售数据
            category_sales = material_category_sales.groupby('物料类别')['销售金额'].sum().reset_index()

            # 合并物料和销售数据
            material_category_roi = pd.merge(material_category_roi, category_sales, on='物料类别', how='left')
            material_category_roi['销售金额'].fillna(0, inplace=True)

            # 计算ROI和覆盖率
            material_category_roi['ROI'] = material_category_roi['销售金额'] / material_category_roi[
                '物料成本'].replace(0, np.nan)
            material_category_roi['ROI'].fillna(0, inplace=True)

            # 计算每种物料类别的经销商覆盖率
            category_coverage = material_data.groupby('物料类别')['经销商名称'].nunique().reset_index()
            total_distributors = material_data['经销商名称'].nunique()
            category_coverage['覆盖率'] = (category_coverage['经销商名称'] / total_distributors * 100).round(2)

            # 合并覆盖率数据
            material_category_roi = pd.merge(material_category_roi, category_coverage[['物料类别', '覆盖率']],
                                             on='物料类别', how='left')

            # 创建物料类别分析图表
            col1, col2 = st.columns(2)

            with col1:
                # 物料类别ROI分析
                category_roi_chart = px.bar(
                    material_category_roi.sort_values('ROI', ascending=False),
                    x='物料类别',
                    y='ROI',
                    title='各物料类别ROI分析',
                    color='ROI',
                    text='ROI',
                    hover_data={
                        '物料成本': ':.2f',
                        '销售金额': ':.2f',
                        'ROI': ':.2f',
                        '覆盖率': ':.2f'
                    }
                )
                category_roi_chart.update_traces(texttemplate='%{text:.2f}', textposition='outside')
                category_roi_chart.update_layout(
                    xaxis_title="物料类别",
                    yaxis_title="ROI"
                )
                st.plotly_chart(category_roi_chart, use_container_width=True)

            with col2:
                # 物料类别覆盖率分析
                coverage_chart = px.bar(
                    material_category_roi.sort_values('覆盖率', ascending=False),
                    x='物料类别',
                    y='覆盖率',
                    title='各物料类别经销商覆盖率',
                    color='覆盖率',
                    text='覆盖率',
                    color_continuous_scale='Blues'
                )
                coverage_chart.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
                coverage_chart.update_layout(
                    xaxis_title="物料类别",
                    yaxis_title="经销商覆盖率(%)"
                )
                st.plotly_chart(coverage_chart, use_container_width=True)

            st.markdown(
                '<div class="highlight-box">左图显示各物料类别的ROI，帮助识别哪些物料类别最具投资回报价值；右图显示各物料类别的经销商覆盖率，帮助了解物料普及程度，高ROI但低覆盖率的物料类别可能存在推广空间。</div>',
                unsafe_allow_html=True)

            # 物料组合分析
            st.markdown('<h3 class="sub-header">物料多样性与ROI关系分析</h3>', unsafe_allow_html=True)

            # 物料多样性与ROI关系
            diversity_roi_chart = px.scatter(
                filtered_distributor,
                x='物料多样性',
                y='ROI',
                color='销售总额',
                size='物料总成本',
                hover_name='经销商名称',
                title='物料多样性与ROI关系',
                hover_data={
                    '物料总成本': ':.2f',
                    '销售总额': ':.2f',
                    'ROI': ':.2f'
                },
                color_continuous_scale='Viridis',
                size_max=40
            )

            diversity_roi_chart.update_layout(
                xaxis_title="使用物料种类数",
                yaxis_title="ROI"
            )

            # 添加趋势线
            diversity_roi_chart.add_trace(
                go.Scatter(
                    x=filtered_distributor['物料多样性'].unique(),
                    y=filtered_distributor.groupby('物料多样性')['ROI'].mean(),
                    mode='lines',
                    name='平均趋势',
                    line=dict(color='red', width=3)
                )
            )

            st.plotly_chart(diversity_roi_chart, use_container_width=True)

            # 物料多样性分组分析
            diversity_bins = [0, 3, 6, 10, 100]
            diversity_labels = ['低多样性(1-3)', '中多样性(4-6)', '高多样性(7-10)', '超高多样性(>10)']

            filtered_distributor['多样性分组'] = pd.cut(
                filtered_distributor['物料多样性'],
                bins=diversity_bins,
                labels=diversity_labels,
                right=False
            )

            diversity_group_analysis = filtered_distributor.groupby('多样性分组').agg({
                '物料总成本': 'mean',
                '销售总额': 'mean',
                'ROI': 'mean',
                '客户代码': 'count'
            }).reset_index()

            diversity_group_analysis.rename(columns={'客户代码': '经销商数量'}, inplace=True)
            diversity_group_analysis['占比'] = (diversity_group_analysis['经销商数量'] / diversity_group_analysis[
                '经销商数量'].sum() * 100).round(2)

            # 格式化数据
            diversity_group_analysis['物料总成本'] = diversity_group_analysis['物料总成本'].round(2)
            diversity_group_analysis['销售总额'] = diversity_group_analysis['销售总额'].round(2)
            diversity_group_analysis['ROI'] = diversity_group_analysis['ROI'].round(2)

            st.markdown('<h4>物料多样性分组分析</h4>', unsafe_allow_html=True)
            st.dataframe(
                diversity_group_analysis,
                use_container_width=True,
                column_config={
                    "物料总成本": st.column_config.NumberColumn("平均物料成本(元)", format="¥%.2f"),
                    "销售总额": st.column_config.NumberColumn("平均销售额(元)", format="¥%.2f"),
                    "ROI": st.column_config.NumberColumn("平均ROI", format="%.2f"),
                    "占比": st.column_config.NumberColumn("占比(%)", format="%.2f%%")
                }
            )

            st.markdown(
                '<div class="highlight-box">上图和表格分析了物料多样性与ROI的关系。通常，使用适当多样化的物料组合(4-10种)的经销商ROI表现更好，这说明多元化的物料组合策略有助于提升销售效果，但过度多样化可能导致资源分散。</div>',
                unsafe_allow_html=True)

            # 物料成本趋势和销售趋势图
            st.markdown('<h3 class="sub-header">物料成本与销售趋势分析</h3>', unsafe_allow_html=True)

            col1, col2 = st.columns(2)

            with col1:
                # 物料成本趋势图
                material_trend = material_data[
                    (material_data['所属区域'].isin(selected_regions)) &
                    (material_data['省份'].isin(selected_provinces))
                    ]

                material_trend = material_trend.groupby(['月份名', '月份'])['物料成本'].sum().reset_index()
                material_trend = material_trend.sort_values('月份')

                material_cost_trend = px.line(
                    material_trend,
                    x='月份名',
                    y='物料成本',
                    title='物料成本月度趋势',
                    markers=True,
                    line_shape='linear'
                )

                material_cost_trend.update_traces(
                    marker=dict(size=10, symbol='circle', line=dict(width=2, color='DarkSlateGrey')),
                    marker_color='rgb(0, 128, 255)',
                    line=dict(width=3),
                    hovertemplate='月份: %{x}<br>物料成本: %{y:.2f}元<extra></extra>'
                )

                material_cost_trend.update_layout(
                    xaxis_title="月份",
                    yaxis_title="物料成本 (元)"
                )

                st.plotly_chart(material_cost_trend, use_container_width=True)

            with col2:
                # 销售趋势图
                sales_trend_data = sales_data[
                    (sales_data['所属区域'].isin(selected_regions)) &
                    (sales_data['省份'].isin(selected_provinces))
                    ]

                sales_trend_data = sales_trend_data.groupby(['月份名', '月份'])['销售金额'].sum().reset_index()
                sales_trend_data = sales_trend_data.sort_values('月份')

                sales_trend_chart = px.line(
                    sales_trend_data,
                    x='月份名',
                    y='销售金额',
                    title='销售金额月度趋势',
                    markers=True,
                    line_shape='linear'
                )

                sales_trend_chart.update_traces(
                    marker=dict(size=10, symbol='circle', line=dict(width=2, color='DarkSlateGrey')),
                    marker_color='rgb(255, 64, 129)',
                    line=dict(width=3),
                    hovertemplate='月份: %{x}<br>销售金额: %{y:.2f}元<extra></extra>'
                )

                sales_trend_chart.update_layout(
                    xaxis_title="月份",
                    yaxis_title="销售金额 (元)"
                )

                st.plotly_chart(sales_trend_chart, use_container_width=True)

            # 前后月份对比
            st.markdown('<h4>月度环比分析</h4>', unsafe_allow_html=True)

            # 按月计算物料成本和销售额
            monthly_data = pd.merge(
                material_trend[['月份名', '月份', '物料成本']],
                sales_trend_data[['月份名', '月份', '销售金额']],
                on=['月份名', '月份'],
                how='outer'
            ).fillna(0)

            monthly_data = monthly_data.sort_values('月份')
            monthly_data['物料成本环比'] = monthly_data['物料成本'].pct_change() * 100
            monthly_data['销售金额环比'] = monthly_data['销售金额'].pct_change() * 100
            monthly_data['ROI'] = monthly_data['销售金额'] / monthly_data['物料成本'].replace(0, np.nan)
            monthly_data['ROI'].fillna(0, inplace=True)
            monthly_data['ROI环比'] = monthly_data['ROI'].pct_change() * 100

            # 格式化数据
            monthly_data['物料成本'] = monthly_data['物料成本'].round(2)
            monthly_data['销售金额'] = monthly_data['销售金额'].round(2)
            monthly_data['ROI'] = monthly_data['ROI'].round(2)
            monthly_data['物料成本环比'] = monthly_data['物料成本环比'].round(2)
            monthly_data['销售金额环比'] = monthly_data['销售金额环比'].round(2)
            monthly_data['ROI环比'] = monthly_data['ROI环比'].round(2)

            # 显示环比数据表格
            st.dataframe(
                monthly_data[1:],  # 去除第一行，因为环比计算会产生NaN
                use_container_width=True,
                column_config={
                    "月份名": "月份",
                    "物料成本": st.column_config.NumberColumn("物料成本(元)", format="¥%.2f"),
                    "销售金额": st.column_config.NumberColumn("销售金额(元)", format="¥%.2f"),
                    "ROI": st.column_config.NumberColumn("ROI", format="%.2f"),
                    "物料成本环比": st.column_config.NumberColumn("物料成本环比(%)", format="%.2f%%"),
                    "销售金额环比": st.column_config.NumberColumn("销售金额环比(%)", format="%.2f%%"),
                    "ROI环比": st.column_config.NumberColumn("ROI环比(%)", format="%.2f%%")
                }
            )

            st.markdown(
                '<div class="highlight-box">上表展示了各月份物料成本、销售金额和ROI的环比变化，帮助您识别业绩变化趋势和物料投入的滞后效应。一般情况下，物料投入的效果在1-2个月内显现，请关注物料成本增长后的销售金额变化模式。</div>',
                unsafe_allow_html=True)

            # 提供下载月度数据的链接
            st.markdown(create_download_link(monthly_data, "月度物料销售数据"), unsafe_allow_html=True)

        # ======= 经销商分析标签页 =======
        with tab3:
            st.markdown(
                '<div class="info-box">本页面分析各经销商的物料使用效率，帮助您识别表现优秀和需要改进的经销商，针对性地提供指导和支持。</div>',
                unsafe_allow_html=True)

            # 经销商绩效图
            st.markdown('<h3 class="sub-header">顶级经销商绩效分析</h3>', unsafe_allow_html=True)

            distributor_performance = filtered_distributor.sort_values('销售总额', ascending=False).head(15)

            distributor_perf_chart = make_subplots(specs=[[{"secondary_y": True}]])

            distributor_perf_chart.add_trace(
                go.Bar(x=distributor_performance['经销商名称'],
                       y=distributor_performance['物料总成本'],
                       name='物料总成本',
                       marker_color='rgba(58, 71, 80, 0.6)',
                       hovertemplate='经销商: %{x}<br>物料总成本: %{y:.2f}元<extra></extra>'),
                secondary_y=False
            )

            distributor_perf_chart.add_trace(
                go.Bar(x=distributor_performance['经销商名称'],
                       y=distributor_performance['销售总额'],
                       name='销售总额',
                       marker_color='rgba(246, 78, 139, 0.6)',
                       hovertemplate='经销商: %{x}<br>销售总额: %{y:.2f}元<extra></extra>'),
                secondary_y=False
            )

            distributor_perf_chart.add_trace(
                go.Scatter(x=distributor_performance['经销商名称'],
                           y=distributor_performance['ROI'],
                           name='ROI',
                           mode='lines+markers+text',
                           line=dict(color='rgb(25, 118, 210)', width=3),
                           marker=dict(size=10),
                           text=distributor_performance['ROI'].round(2),
                           textposition='top center',
                           hovertemplate='经销商: %{x}<br>ROI: %{y:.2f}<extra></extra>'),
                secondary_y=True
            )

            distributor_perf_chart.update_layout(
                title_text='销售额Top 15的经销商绩效',
                barmode='group',
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                )
            )

            distributor_perf_chart.update_xaxes(tickangle=-45)
            distributor_perf_chart.update_yaxes(title_text='金额 (元)', secondary_y=False)
            distributor_perf_chart.update_yaxes(title_text='ROI', secondary_y=True)

            st.plotly_chart(distributor_perf_chart, use_container_width=True)
            st.markdown(
                '<div class="highlight-box">此图表展示销售额前15的经销商的物料投入和销售产出情况。请注意比较物料成本与销售额的匹配程度，以及各经销商的ROI表现。部分销售额高但ROI较低的经销商可能存在物料使用效率问题。</div>',
                unsafe_allow_html=True)

            # 经销商物料使用详情
            st.markdown('<h3 class="sub-header">经销商物料使用详情</h3>', unsafe_allow_html=True)

            # 选择经销商
            distributor_list = sorted(filtered_distributor['经销商名称'].unique())
            if len(distributor_list) > 0:
                selected_distributor = st.selectbox("选择经销商查看详情:", distributor_list)

                # 获取选中经销商的详细信息
                dist_details = filtered_distributor[filtered_distributor['经销商名称'] == selected_distributor].iloc[0]
                dist_code = dist_details['客户代码']

                # 显示经销商基础信息
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("物料总成本", format_currency(dist_details['物料总成本']))
                with col2:
                    st.metric("销售总额", format_currency(dist_details['销售总额']))
                with col3:
                    roi_color = "success-metric" if dist_details['ROI'] >= 2.0 else "warning-metric" if dist_details[
                                                                                                            'ROI'] >= 1.0 else "danger-metric"
                    st.markdown(
                        f'<div class="metric-card"><h3 class="{roi_color}">ROI: {dist_details["ROI"]:.2f}</h3></div>',
                        unsafe_allow_html=True)
                with col4:
                    st.metric("物料多样性", f"{dist_details['物料多样性']:.0f}种")

                # 获取该经销商的物料使用详情
                dist_materials = filtered_material[filtered_material['客户代码'] == dist_code]
                dist_materials_summary = dist_materials.groupby(['产品名称', '物料类别']).agg({
                    '求和项:数量（箱）': 'sum',
                    '物料成本': 'sum'
                }).reset_index()

                # 按物料成本排序
                dist_materials_summary = dist_materials_summary.sort_values('物料成本', ascending=False)

                # 创建物料使用饼图
                col1, col2 = st.columns(2)

                with col1:
                    # 物料类别占比饼图
                    category_pie = px.pie(
                        dist_materials_summary,
                        values='物料成本',
                        names='物料类别',
                        title=f"{selected_distributor}的物料成本类别占比",
                        hover_data=['物料成本'],
                        labels={'物料成本': '物料成本(元)'}
                    )

                    category_pie.update_traces(textposition='inside', textinfo='percent+label')
                    category_pie.update_layout(uniformtext_minsize=12, uniformtext_mode='hide')

                    st.plotly_chart(category_pie, use_container_width=True)

                with col2:
                    # 物料使用分布图
                    materials_bar = px.bar(
                        dist_materials_summary.head(10),
                        x='产品名称',
                        y='物料成本',
                        color='物料类别',
                        title=f"{selected_distributor}的物料成本分布(Top 10)",
                        text='物料成本'
                    )

                    materials_bar.update_traces(texttemplate='%{text:.2f}', textposition='outside')
                    materials_bar.update_layout(
                        xaxis_tickangle=-45,
                        xaxis_title="物料名称",
                        yaxis_title="物料成本(元)"
                    )

                    st.plotly_chart(materials_bar, use_container_width=True)

                # 显示该经销商的物料详情表格
                st.markdown('<h4>物料使用详情</h4>', unsafe_allow_html=True)

                # 格式化数据
                dist_materials_summary['求和项:数量（箱）'] = dist_materials_summary['求和项:数量（箱）'].round(2)
                dist_materials_summary['物料成本'] = dist_materials_summary['物料成本'].round(2)
                dist_materials_summary['成本占比'] = (
                            dist_materials_summary['物料成本'] / dist_materials_summary['物料成本'].sum() * 100).round(
                    2)

                st.dataframe(
                    dist_materials_summary,
                    use_container_width=True,
                    column_config={
                        "产品名称": "物料名称",
                        "求和项:数量（箱）": st.column_config.NumberColumn("数量(箱)", format="%.2f"),
                        "物料成本": st.column_config.NumberColumn("物料成本(元)", format="¥%.2f"),
                        "成本占比": st.column_config.NumberColumn("成本占比(%)", format="%.2f%%")
                    }
                )

                # 获取类似经销商推荐
                st.markdown('<h4>类似经销商对比</h4>', unsafe_allow_html=True)

                # 计算经销商相似度（基于物料成本规模）
                cost_mean = filtered_distributor['物料总成本'].mean()
                cost_std = filtered_distributor['物料总成本'].std()

                similar_dist = filtered_distributor[
                    (filtered_distributor['经销商名称'] != selected_distributor) &
                    (filtered_distributor['物料总成本'] > dist_details['物料总成本'] * 0.8) &
                    (filtered_distributor['物料总成本'] < dist_details['物料总成本'] * 1.2)
                    ]

                if len(similar_dist) > 0:
                    # 取Top 5类似经销商
                    similar_dist = similar_dist.sort_values('ROI', ascending=False).head(5)

                    # 创建对比表格
                    comparison_data = pd.concat([
                        pd.DataFrame({
                            '经销商名称': [selected_distributor],
                            '物料总成本': [dist_details['物料总成本']],
                            '销售总额': [dist_details['销售总额']],
                            'ROI': [dist_details['ROI']],
                            '物料多样性': [dist_details['物料多样性']],
                            '物料销售比率': [dist_details['物料销售比率']]
                        }),
                        similar_dist[['经销商名称', '物料总成本', '销售总额', 'ROI', '物料多样性', '物料销售比率']]
                    ])

                    # 格式化数据
                    comparison_data['物料总成本'] = comparison_data['物料总成本'].round(2)
                    comparison_data['销售总额'] = comparison_data['销售总额'].round(2)
                    comparison_data['ROI'] = comparison_data['ROI'].round(2)
                    comparison_data['物料多样性'] = comparison_data['物料多样性'].round(0)
                    comparison_data['物料销售比率'] = comparison_data['物料销售比率'].round(2)

                    st.dataframe(
                        comparison_data,
                        use_container_width=True,
                        column_config={
                            "物料总成本": st.column_config.NumberColumn("物料总成本(元)", format="¥%.2f"),
                            "销售总额": st.column_config.NumberColumn("销售总额(元)", format="¥%.2f"),
                            "ROI": st.column_config.NumberColumn("ROI", format="%.2f"),
                            "物料多样性": st.column_config.NumberColumn("物料多样性(种)", format="%d"),
                            "物料销售比率": st.column_config.NumberColumn("物料销售比率(%)", format="%.2f%%")
                        }
                    )

                    st.markdown(
                        '<div class="highlight-box">此表对比了与所选经销商物料成本规模相近的其他经销商表现。通过对比，您可以发现类似规模经销商的最佳实践，为所选经销商提供改进建议。</div>',
                        unsafe_allow_html=True)
                else:
                    st.write("未找到规模相似的经销商进行对比")
            else:
                st.warning("没有可用的经销商数据")

            # 高效和低效经销商表格
            st.markdown('<h3 class="sub-header">经销商效率分析</h3>', unsafe_allow_html=True)

            col1, col2 = st.columns(2)

            with col1:
                st.markdown('<h4>高效物料投放经销商 Top 10</h4>', unsafe_allow_html=True)
                st.write("以下经销商在物料使用上表现优异，ROI值较高，可作为标杆学习。")

                efficient_distributors = filtered_distributor.sort_values('ROI', ascending=False).head(10)

                # 格式化数据
                efficient_display = efficient_distributors[
                    ['经销商名称', '物料总成本', '销售总额', 'ROI', '物料销售比率']].copy()
                efficient_display['物料总成本'] = efficient_display['物料总成本'].round(2)
                efficient_display['销售总额'] = efficient_display['销售总额'].round(2)
                efficient_display['ROI'] = efficient_display['ROI'].round(2)
                efficient_display['物料销售比率'] = efficient_display['物料销售比率'].round(2)

                st.dataframe(
                    efficient_display,
                    use_container_width=True,
                    column_config={
                        "物料总成本": st.column_config.NumberColumn("物料总成本(元)", format="¥%.2f"),
                        "销售总额": st.column_config.NumberColumn("销售总额(元)", format="¥%.2f"),
                        "ROI": st.column_config.NumberColumn("ROI", format="%.2f"),
                        "物料销售比率": st.column_config.NumberColumn("物料销售比率(%)", format="%.2f%%")
                    }
                )

            with col2:
                st.markdown('<h4>待优化物料投放经销商 Top 10</h4>', unsafe_allow_html=True)
                st.write("以下经销商在物料使用上有改进空间，ROI值较低，需要提供针对性指导。")

                # 只考虑物料成本大于0的经销商
                inefficient_distributors = filtered_distributor[
                    (filtered_distributor['物料总成本'] > 0) &
                    (filtered_distributor['销售总额'] > 0)
                    ].sort_values('ROI').head(10)

                # 格式化数据
                inefficient_display = inefficient_distributors[
                    ['经销商名称', '物料总成本', '销售总额', 'ROI', '物料销售比率']].copy()
                inefficient_display['物料总成本'] = inefficient_display['物料总成本'].round(2)
                inefficient_display['销售总额'] = inefficient_display['销售总额'].round(2)
                inefficient_display['ROI'] = inefficient_display['ROI'].round(2)
                inefficient_display['物料销售比率'] = inefficient_display['物料销售比率'].round(2)

                st.dataframe(
                    inefficient_display,
                    use_container_width=True,
                    column_config={
                        "物料总成本": st.column_config.NumberColumn("物料总成本(元)", format="¥%.2f"),
                        "销售总额": st.column_config.NumberColumn("销售总额(元)", format="¥%.2f"),
                        "ROI": st.column_config.NumberColumn("ROI", format="%.2f"),
                        "物料销售比率": st.column_config.NumberColumn("物料销售比率(%)", format="%.2f%%")
                    }
                )

            # 经销商表现分布
            st.markdown('<h3 class="sub-header">经销商ROI分布分析</h3>', unsafe_allow_html=True)

            # 创建ROI分布直方图
            roi_hist = px.histogram(
                filtered_distributor[filtered_distributor['ROI'] > 0],
                x='ROI',
                title='经销商ROI分布',
                nbins=20,
                histnorm='percent',
                marginal='box',
                color_discrete_sequence=['#1E88E5']
            )

            # 添加参考线
            roi_hist.add_shape(
                type="line",
                x0=1, y0=0,
                x1=1, y1=roi_hist.data[0].y.max(),
                line=dict(color="red", width=2, dash="dash")
            )

            roi_hist.add_shape(
                type="line",
                x0=2, y0=0,
                x1=2, y1=roi_hist.data[0].y.max(),
                line=dict(color="green", width=2, dash="dash")
            )

            roi_hist.add_annotation(
                x=1, y=roi_hist.data[0].y.max() * 0.9,
                text="ROI=1",
                showarrow=False,
                font=dict(color="red")
            )

            roi_hist.add_annotation(
                x=2, y=roi_hist.data[0].y.max() * 0.9,
                text="ROI=2",
                showarrow=False,
                font=dict(color="green")
            )

            roi_hist.update_layout(
                xaxis_title="ROI值",
                yaxis_title="占比(%)",
                bargap=0.1
            )

            st.plotly_chart(roi_hist, use_container_width=True)

            # 经销商ROI分组统计
            roi_bins = [0, 1, 2, 3, 100]
            roi_labels = ['低效(ROI<1)', '一般(1≤ROI<2)', '良好(2≤ROI<3)', '优秀(ROI≥3)']

            filtered_distributor['ROI分组'] = pd.cut(
                filtered_distributor['ROI'],
                bins=roi_bins,
                labels=roi_labels,
                right=False
            )

            roi_group_analysis = filtered_distributor.groupby('ROI分组').agg({
                '物料总成本': 'sum',
                '销售总额': 'sum',
                '客户代码': 'count'
            }).reset_index()

            roi_group_analysis.rename(columns={'客户代码': '经销商数量'}, inplace=True)
            roi_group_analysis['占比'] = (
                        roi_group_analysis['经销商数量'] / roi_group_analysis['经销商数量'].sum() * 100).round(2)
            roi_group_analysis['平均物料成本'] = (
                        roi_group_analysis['物料总成本'] / roi_group_analysis['经销商数量']).round(2)
            roi_group_analysis['平均销售额'] = (
                        roi_group_analysis['销售总额'] / roi_group_analysis['经销商数量']).round(2)

            # 格式化数据
            roi_group_analysis['物料总成本'] = roi_group_analysis['物料总成本'].round(2)
            roi_group_analysis['销售总额'] = roi_group_analysis['销售总额'].round(2)

            st.markdown('<h4>经销商ROI分组统计</h4>', unsafe_allow_html=True)
            st.dataframe(
                roi_group_analysis[
                    ['ROI分组', '经销商数量', '占比', '物料总成本', '销售总额', '平均物料成本', '平均销售额']],
                use_container_width=True,
                column_config={
                    "物料总成本": st.column_config.NumberColumn("物料总成本(元)", format="¥%.2f"),
                    "销售总额": st.column_config.NumberColumn("销售总额(元)", format="¥%.2f"),
                    "平均物料成本": st.column_config.NumberColumn("平均物料成本(元)", format="¥%.2f"),
                    "平均销售额": st.column_config.NumberColumn("平均销售额(元)", format="¥%.2f"),
                    "占比": st.column_config.NumberColumn("占比(%)", format="%.2f%%")
                }
            )

            st.markdown(
                '<div class="highlight-box">经销商ROI分布图和分组统计表展示了经销商ROI的整体分布情况。理想情况下，大部分经销商应集中在ROI≥2的区间。如果低效经销商占比过高，说明物料使用效率普遍不足，需要加强物料使用培训和指导。</div>',
                unsafe_allow_html=True)

            # 提供下载经销商数据的链接
            st.markdown(create_download_link(filtered_distributor, "经销商物料销售数据"), unsafe_allow_html=True)

        # ======= ROI分析标签页 =======
        with tab4:
            st.markdown(
                '<div class="info-box">本页面深入分析各物料的投资回报率(ROI)，帮助您识别效果最佳和效果最差的物料类型，优化物料投放组合。</div>',
                unsafe_allow_html=True)

            # 物料ROI分析
            st.markdown('<h3 class="sub-header">物料ROI分析</h3>', unsafe_allow_html=True)

            # 物料ROI图表
            material_roi = filtered_material.groupby(['产品代码', '产品名称', '物料类别']).agg({
                '物料成本': 'sum',
                '求和项:数量（箱）': 'sum'
            }).reset_index()

            # 匹配销售数据(通过产品代码前缀)
            sales_by_product = filtered_sales.copy()
            sales_by_product['匹配代码'] = sales_by_product['产品代码'].str[:6]
            material_roi['匹配代码'] = material_roi['产品代码'].str[:6]

            # 按匹配代码汇总销售数据
            sales_summary = sales_by_product.groupby('匹配代码')['销售金额'].sum().reset_index()

            # 合并物料和销售数据
            material_roi = pd.merge(material_roi, sales_summary, on='匹配代码', how='left')
            material_roi['销售金额'].fillna(0, inplace=True)

            # 计算ROI
            material_roi['ROI'] = material_roi['销售金额'] / material_roi['物料成本'].replace(0, np.nan)
            material_roi['ROI'].fillna(0, inplace=True)

            # 计算物料覆盖率(经销商数量)
            material_coverage = filtered_material.groupby('产品代码')['经销商名称'].nunique().reset_index()
            material_coverage.rename(columns={'经销商名称': '经销商数量'}, inplace=True)

            # 计算覆盖率百分比
            total_distributors = filtered_material['经销商名称'].nunique()
            material_coverage['覆盖率'] = (material_coverage['经销商数量'] / total_distributors * 100).round(2)

            # 合并覆盖率数据
            material_roi = pd.merge(material_roi, material_coverage, on='产品代码', how='left')

            # 只保留ROI > 0和物料成本 > 0的物料
            material_roi = material_roi[
                (material_roi['ROI'] > 0) &
                (material_roi['物料成本'] > 0)
                ].sort_values('ROI', ascending=False)

            col1, col2 = st.columns(2)

            with col1:
                # ROI最高的物料
                material_roi_chart = px.bar(
                    material_roi.head(15),
                    x='产品名称',
                    y='ROI',
                    title='ROI最高的物料Top 15',
                    color='物料类别',
                    text='ROI',
                    hover_data={
                        '物料成本': ':.2f',
                        '销售金额': ':.2f',
                        '覆盖率': ':.2f',
                        '经销商数量': True
                    }
                )
                material_roi_chart.update_traces(texttemplate='%{text:.2f}', textposition='outside')
                material_roi_chart.update_layout(
                    xaxis_tickangle=-45,
                    xaxis_title="物料名称",
                    yaxis_title="ROI"
                )

                st.plotly_chart(material_roi_chart, use_container_width=True)

            with col2:
                # 覆盖率最高的物料
                coverage_chart = px.bar(
                    material_roi.sort_values('覆盖率', ascending=False).head(15),
                    x='产品名称',
                    y='覆盖率',
                    title='覆盖率最高的物料Top 15',
                    color='物料类别',
                    text='覆盖率',
                    hover_data={
                        '物料成本': ':.2f',
                        '销售金额': ':.2f',
                        'ROI': ':.2f',
                        '经销商数量': True
                    }
                )
                coverage_chart.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
                coverage_chart.update_layout(
                    xaxis_tickangle=-45,
                    xaxis_title="物料名称",
                    yaxis_title="经销商覆盖率(%)"
                )

                st.plotly_chart(coverage_chart, use_container_width=True)

            # 物料ROI与覆盖率散点图
            roi_coverage_scatter = px.scatter(
                material_roi,
                x='ROI',
                y='覆盖率',
                color='物料类别',
                size='物料成本',
                hover_name='产品名称',
                title='物料ROI与覆盖率关系',
                size_max=40,
                hover_data={
                    '物料成本': ':.2f',
                    '销售金额': ':.2f',
                    '经销商数量': True
                }
            )

            roi_coverage_scatter.update_layout(
                xaxis_title="ROI",
                yaxis_title="经销商覆盖率(%)",
                xaxis=dict(
                    range=[0, material_roi['ROI'].quantile(0.95)]  # 限制x轴范围，避免极端值影响视觉效果
                )
            )

            # 添加参考线 - ROI=1
            roi_coverage_scatter.add_shape(
                type="line",
                x0=1, y0=0,
                x1=1, y1=100,
                line=dict(color="red", width=2, dash="dash")
            )

            # 添加参考线 - ROI=2
            roi_coverage_scatter.add_shape(
                type="line",
                x0=2, y0=0,
                x1=2, y1=100,
                line=dict(color="green", width=2, dash="dash")
            )

            st.plotly_chart(roi_coverage_scatter, use_container_width=True)

            st.markdown(
                '<div class="highlight-box">上图展示了各物料的ROI与覆盖率关系，点的大小表示物料成本。位于图右上方的物料具有高ROI和高覆盖率，是理想的投放物料；位于左上方的物料覆盖率高但ROI低，需要优化使用方式；位于右下方的物料ROI高但覆盖率低，可考虑扩大投放范围。</div>',
                unsafe_allow_html=True)

            # 物料ROI详情表格
            st.markdown('<h3 class="sub-header">物料投资回报详情</h3>', unsafe_allow_html=True)

            # 添加物料投入产出比率
            material_roi['物料投入产出比'] = (material_roi['物料成本'] / material_roi['销售金额'] * 100).round(2)

            # 添加单箱物料成本和销售贡献
            material_roi['单箱成本'] = (material_roi['物料成本'] / material_roi['求和项:数量（箱）']).round(2)
            material_roi['单箱销售贡献'] = (material_roi['销售金额'] / material_roi['求和项:数量（箱）']).round(2)

            # 格式化数据
            material_roi_display = material_roi.copy()
            material_roi_display['物料成本'] = material_roi_display['物料成本'].round(2)
            material_roi_display['销售金额'] = material_roi_display['销售金额'].round(2)
            material_roi_display['ROI'] = material_roi_display['ROI'].round(2)
            material_roi_display['求和项:数量（箱）'] = material_roi_display['求和项:数量（箱）'].round(2)

            # 表现评估
            material_roi_display['表现评估'] = np.where(material_roi_display['ROI'] >= 2.0, '优秀',
                                                        np.where(material_roi_display['ROI'] >= 1.0, '良好', '需改进'))

            # 显示物料ROI详情表
            st.dataframe(
                material_roi_display[[
                    '产品名称', '物料类别', '物料成本', '销售金额', 'ROI', '覆盖率',
                    '经销商数量', '求和项:数量（箱）', '单箱成本', '单箱销售贡献', '表现评估'
                ]],
                use_container_width=True,
                column_config={
                    "物料成本": st.column_config.NumberColumn("物料成本(元)", format="¥%.2f"),
                    "销售金额": st.column_config.NumberColumn("销售金额(元)", format="¥%.2f"),
                    "ROI": st.column_config.NumberColumn("ROI", format="%.2f"),
                    "覆盖率": st.column_config.NumberColumn("覆盖率(%)", format="%.2f%%"),
                    "求和项:数量（箱）": st.column_config.NumberColumn("数量(箱)", format="%.2f"),
                    "单箱成本": st.column_config.NumberColumn("单箱成本(元)", format="¥%.2f"),
                    "单箱销售贡献": st.column_config.NumberColumn("单箱销售贡献(元)", format="¥%.2f")
                }
            )

            # 物料类别ROI分析
            st.markdown('<h3 class="sub-header">物料类别ROI分析</h3>', unsafe_allow_html=True)

            # 按物料类别汇总数据
            category_roi = material_roi.groupby('物料类别').agg({
                '物料成本': 'sum',
                '销售金额': 'sum',
                '产品代码': 'count',
                '经销商数量': 'mean'
            }).reset_index()

            category_roi.rename(columns={'产品代码': '物料数量'}, inplace=True)
            category_roi['ROI'] = category_roi['销售金额'] / category_roi['物料成本']
            category_roi['平均物料成本'] = category_roi['物料成本'] / category_roi['物料数量']

            # 计算平均覆盖率
            category_roi['平均覆盖率'] = (material_roi.groupby('物料类别')['覆盖率'].mean()).values

            # 格式化数据
            category_roi['物料成本'] = category_roi['物料成本'].round(2)
            category_roi['销售金额'] = category_roi['销售金额'].round(2)
            category_roi['ROI'] = category_roi['ROI'].round(2)
            category_roi['平均物料成本'] = category_roi['平均物料成本'].round(2)
            category_roi['平均覆盖率'] = category_roi['平均覆盖率'].round(2)
            category_roi['经销商数量'] = category_roi['经销商数量'].round(0)

            # 创建类别ROI柱状图
            category_bar = px.bar(
                category_roi.sort_values('ROI', ascending=False),
                x='物料类别',
                y='ROI',
                color='物料类别',
                title='各物料类别ROI比较',
                text='ROI',
                hover_data={
                    '物料成本': ':.2f',
                    '销售金额': ':.2f',
                    '物料数量': True,
                    '平均覆盖率': ':.2f'
                }
            )

            category_bar.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            category_bar.update_layout(
                xaxis_title="物料类别",
                yaxis_title="ROI"
            )

            # 添加参考线 - ROI=1
            category_bar.add_shape(
                type="line",
                x0=-0.5, y0=1,
                x1=len(category_roi) - 0.5, y1=1,
                line=dict(color="red", width=2, dash="dash")
            )

            # 添加参考线 - ROI=2
            category_bar.add_shape(
                type="line",
                x0=-0.5, y0=2,
                x1=len(category_roi) - 0.5, y1=2,
                line=dict(color="green", width=2, dash="dash")
            )

            st.plotly_chart(category_bar, use_container_width=True)

            # 显示类别ROI详情表
            st.dataframe(
                category_roi.sort_values('ROI', ascending=False),
                use_container_width=True,
                column_config={
                    "物料成本": st.column_config.NumberColumn("物料成本(元)", format="¥%.2f"),
                    "销售金额": st.column_config.NumberColumn("销售金额(元)", format="¥%.2f"),
                    "ROI": st.column_config.NumberColumn("ROI", format="%.2f"),
                    "平均物料成本": st.column_config.NumberColumn("平均物料成本(元)", format="¥%.2f"),
                    "平均覆盖率": st.column_config.NumberColumn("平均覆盖率(%)", format="%.2f%%"),
                    "经销商数量": st.column_config.NumberColumn("平均经销商数量", format="%.0f")
                }
            )

            st.markdown(
                '<div class="highlight-box">各物料类别ROI比较图和表格展示了不同物料类别的投资回报表现。ROI高的物料类别应增加投放比例，ROI低的物料类别应减少投放或改进使用方式。同时结合平均覆盖率，评估各物料类别的使用普及情况。</div>',
                unsafe_allow_html=True)

            # 提供下载物料ROI数据的链接
            st.markdown(create_download_link(material_roi_display, "物料ROI分析数据"), unsafe_allow_html=True)

            # ======= 客户价值分层标签页 =======
        with tab5:
            st.markdown(
                '<div class="info-box">本页面根据物料ROI和销售额对经销商进行价值分层，帮助您识别不同价值的客户群体，采取差异化管理策略。</div>',
                unsafe_allow_html=True)

            # 客户价值分层分析
            st.markdown('<h3 class="sub-header">客户价值分层分布</h3>', unsafe_allow_html=True)

            # 计算各分层客户数量
            segment_counts = filtered_distributor['客户价值分层'].value_counts().reset_index()
            segment_counts.columns = ['客户价值分层', '经销商数量']
            segment_counts['占比'] = (segment_counts['经销商数量'] / segment_counts['经销商数量'].sum() * 100).round(2)

            # 为价值分层设置颜色
            segment_colors = {
                '高价值客户': '#4CAF50',
                '成长型客户': '#2196F3',
                '稳定型客户': '#FFC107',
                '低效型客户': '#F44336'
            }

            # 价值分层饼图
            col1, col2 = st.columns(2)

            with col1:
                segment_pie = px.pie(
                    segment_counts,
                    values='经销商数量',
                    names='客户价值分层',
                    title='客户价值分层分布',
                    color='客户价值分层',
                    color_discrete_map=segment_colors,
                    hole=0.4
                )

                segment_pie.update_traces(textposition='inside', textinfo='percent+label')
                segment_pie.update_layout(
                    legend_title="客户价值分层",
                    annotations=[dict(text='客户分布', x=0.5, y=0.5, font_size=20, showarrow=False)]
                )

                st.plotly_chart(segment_pie, use_container_width=True)

            with col2:
                # 价值分层柱状图（按经销商数量）
                segment_bar = px.bar(
                    segment_counts,
                    x='客户价值分层',
                    y='经销商数量',
                    title='各价值分层经销商数量',
                    text='经销商数量',
                    color='客户价值分层',
                    color_discrete_map=segment_colors
                )

                segment_bar.update_traces(texttemplate='%{text}', textposition='outside')
                segment_bar.update_layout(
                    xaxis_title="客户价值分层",
                    yaxis_title="经销商数量"
                )

                st.plotly_chart(segment_bar, use_container_width=True)

            # 价值分层详细解释
            st.markdown('<h4>客户价值分层说明</h4>', unsafe_allow_html=True)

            segment_desc = pd.DataFrame({
                '客户价值分层': ['高价值客户', '成长型客户', '稳定型客户', '低效型客户'],
                '定义': [
                    'ROI ≥ 2.0且销售额高于75%分位数的经销商',
                    'ROI ≥ 1.0且销售额高于中位数的经销商',
                    'ROI ≥ 1.0的其他经销商',
                    'ROI < 1.0的经销商'
                ],
                '特点': [
                    '物料使用效率高，销售表现优异，是核心价值客户',
                    '物料使用有效，销售潜力大，是重点培养对象',
                    '物料投入产出平衡，销售表现稳定',
                    '物料投入未产生有效回报，需要改进'
                ],
                '管理策略': [
                    '维护关系，优先资源配置，挖掘最佳实践',
                    '加大支持力度，培养成为高价值客户',
                    '保持稳定供应，提升物料使用效率',
                    '诊断问题，提供培训，调整物料投放策略'
                ]
            })

            st.dataframe(segment_desc, use_container_width=True)

            # 各分层客户的物料与销售情况
            st.markdown('<h3 class="sub-header">各价值分层客户表现对比</h3>', unsafe_allow_html=True)

            # 计算各分层汇总数据
            segment_summary = filtered_distributor.groupby('客户价值分层').agg({
                '物料总成本': 'sum',
                '销售总额': 'sum',
                '客户代码': 'count',
                'ROI': 'mean',
                '物料多样性': 'mean',
                '物料销售比率': 'mean'
            }).reset_index()

            segment_summary.rename(columns={'客户代码': '经销商数量'}, inplace=True)
            segment_summary['平均物料成本'] = segment_summary['物料总成本'] / segment_summary['经销商数量']
            segment_summary['平均销售额'] = segment_summary['销售总额'] / segment_summary['经销商数量']
            segment_summary['物料成本占比'] = (
                        segment_summary['物料总成本'] / segment_summary['物料总成本'].sum() * 100).round(2)
            segment_summary['销售额占比'] = (
                        segment_summary['销售总额'] / segment_summary['销售总额'].sum() * 100).round(2)

            # 格式化数据
            for col in ['物料总成本', '销售总额', '平均物料成本', '平均销售额', 'ROI', '物料多样性', '物料销售比率']:
                segment_summary[col] = segment_summary[col].round(2)

            # 创建多指标对比图
            segment_metrics = make_subplots(
                rows=2, cols=2,
                subplot_titles=("各分层物料总成本占比", "各分层销售总额占比", "各分层平均ROI", "各分层平均物料多样性"),
                specs=[[{"type": "pie"}, {"type": "pie"}], [{"type": "bar"}, {"type": "bar"}]]
            )

            # 物料成本占比饼图
            segment_metrics.add_trace(
                go.Pie(
                    labels=segment_summary['客户价值分层'],
                    values=segment_summary['物料总成本'],
                    name="物料总成本",
                    marker_colors=[segment_colors.get(seg, '#000000') for seg in segment_summary['客户价值分层']],
                    textinfo='percent+label'
                ),
                row=1, col=1
            )

            # 销售额占比饼图
            segment_metrics.add_trace(
                go.Pie(
                    labels=segment_summary['客户价值分层'],
                    values=segment_summary['销售总额'],
                    name="销售总额",
                    marker_colors=[segment_colors.get(seg, '#000000') for seg in segment_summary['客户价值分层']],
                    textinfo='percent+label'
                ),
                row=1, col=2
            )

            # 平均ROI柱状图
            segment_metrics.add_trace(
                go.Bar(
                    x=segment_summary['客户价值分层'],
                    y=segment_summary['ROI'],
                    name="平均ROI",
                    marker_color=[segment_colors.get(seg, '#000000') for seg in segment_summary['客户价值分层']],
                    text=segment_summary['ROI']
                ),
                row=2, col=1
            )

            # 平均物料多样性柱状图
            segment_metrics.add_trace(
                go.Bar(
                    x=segment_summary['客户价值分层'],
                    y=segment_summary['物料多样性'],
                    name="平均物料多样性",
                    marker_color=[segment_colors.get(seg, '#000000') for seg in segment_summary['客户价值分层']],
                    text=segment_summary['物料多样性']
                ),
                row=2, col=2
            )

            segment_metrics.update_layout(
                title_text="各价值分层客户表现多维对比",
                height=800
            )

            # 更新柱状图文本位置
            segment_metrics.update_traces(
                texttemplate='%{text:.2f}',
                textposition='outside',
                row=2, col=1
            )

            segment_metrics.update_traces(
                texttemplate='%{text:.2f}',
                textposition='outside',
                row=2, col=2
            )

            st.plotly_chart(segment_metrics, use_container_width=True)

            # 显示各分层汇总数据表格
            st.markdown('<h4>各价值分层客户汇总数据</h4>', unsafe_allow_html=True)
            st.dataframe(
                segment_summary,
                use_container_width=True,
                column_config={
                    "物料总成本": st.column_config.NumberColumn("物料总成本(元)", format="¥%.2f"),
                    "销售总额": st.column_config.NumberColumn("销售总额(元)", format="¥%.2f"),
                    "ROI": st.column_config.NumberColumn("平均ROI", format="%.2f"),
                    "平均物料成本": st.column_config.NumberColumn("平均物料成本(元)", format="¥%.2f"),
                    "平均销售额": st.column_config.NumberColumn("平均销售额(元)", format="¥%.2f"),
                    "物料销售比率": st.column_config.NumberColumn("平均物料销售比率(%)", format="%.2f%%"),
                    "物料成本占比": st.column_config.NumberColumn("物料成本占比(%)", format="%.2f%%"),
                    "销售额占比": st.column_config.NumberColumn("销售额占比(%)", format="%.2f%%")
                }
            )

            st.markdown(
                '<div class="highlight-box">从上述分析可以看出，各价值分层的经销商在物料使用效率和销售表现上存在显著差异。通常高价值客户和成长型客户贡献了大部分销售额，且物料使用效率较高；而低效型客户虽占用了部分物料资源，但销售回报较低。针对不同价值分层采取差异化管理策略，有助于提高整体物料投入产出效率。</div>',
                unsafe_allow_html=True)

            # 物料使用模式差异分析
            st.markdown('<h3 class="sub-header">各价值分层物料使用模式分析</h3>', unsafe_allow_html=True)

            # 按客户价值分层和物料类别分析
            segment_material_usage = filtered_material.copy()

            # 添加客户价值分层信息
            segment_material_usage = pd.merge(
                segment_material_usage,
                filtered_distributor[['客户代码', '客户价值分层']],
                on='客户代码',
                how='left'
            )

            # 按价值分层和物料类别汇总物料成本
            segment_category_cost = segment_material_usage.groupby(['客户价值分层', '物料类别'])[
                '物料成本'].sum().reset_index()

            # 计算各分层物料类别占比
            segment_total_cost = segment_material_usage.groupby('客户价值分层')['物料成本'].sum().reset_index()
            segment_total_cost.rename(columns={'物料成本': '分层总成本'}, inplace=True)

            segment_category_cost = pd.merge(
                segment_category_cost,
                segment_total_cost,
                on='客户价值分层',
                how='left'
            )

            segment_category_cost['占比'] = (
                        segment_category_cost['物料成本'] / segment_category_cost['分层总成本'] * 100).round(2)

            # 创建分层物料使用对比图
            segment_material_chart = px.bar(
                segment_category_cost,
                x='客户价值分层',
                y='占比',
                color='物料类别',
                title='各价值分层物料类别使用占比',
                text='占比',
                barmode='stack'
            )

            segment_material_chart.update_traces(texttemplate='%{text:.1f}%', textposition='inside')
            segment_material_chart.update_layout(
                xaxis_title="客户价值分层",
                yaxis_title="物料类别占比(%)"
            )

            st.plotly_chart(segment_material_chart, use_container_width=True)

            # 计算物料多样性分布
            segment_diversity = filtered_distributor.groupby(['客户价值分层', '物料多样性']).size().reset_index(
                name='经销商数量')

            # 创建分层物料多样性直方图
            segment_diversity_chart = px.histogram(
                filtered_distributor,
                x='物料多样性',
                color='客户价值分层',
                title='各价值分层物料多样性分布',
                opacity=0.7,
                nbins=20,
                color_discrete_map=segment_colors,
                barmode='overlay',
                histnorm='probability',
                marginal='box'
            )

            segment_diversity_chart.update_layout(
                xaxis_title="物料多样性(种)",
                yaxis_title="概率密度"
            )

            st.plotly_chart(segment_diversity_chart, use_container_width=True)

            st.markdown(
                '<div class="highlight-box">各价值分层客户的物料使用模式存在明显差异。通过对比物料类别占比和多样性分布可以发现，高价值客户通常使用更多样化的物料组合，且对某些特定物料类别的偏好更明显。这些洞察可以指导我们为不同价值层次的客户提供更有针对性的物料组合建议。</div>',
                unsafe_allow_html=True)

            # 各价值分层客户名单
            st.markdown('<h3 class="sub-header">各价值分层客户名单</h3>', unsafe_allow_html=True)

            # 创建分层客户数据透视表
            segment_customer_pivot = filtered_distributor.pivot_table(
                index='经销商名称',
                columns='客户价值分层',
                values='ROI',
                aggfunc='first'
            ).fillna(0)

            # 添加选项卡显示各分层客户名单
            segment_tabs = st.tabs(segment_summary['客户价值分层'].tolist())

            for i, segment_name in enumerate(segment_summary['客户价值分层']):
                with segment_tabs[i]:
                    segment_customers = filtered_distributor[
                        filtered_distributor['客户价值分层'] == segment_name].sort_values('销售总额', ascending=False)

                    # 格式化数据
                    segment_customers_display = segment_customers[
                        ['经销商名称', '物料总成本', '销售总额', 'ROI', '物料多样性', '物料销售比率']].copy()
                    segment_customers_display['物料总成本'] = segment_customers_display['物料总成本'].round(2)
                    segment_customers_display['销售总额'] = segment_customers_display['销售总额'].round(2)
                    segment_customers_display['ROI'] = segment_customers_display['ROI'].round(2)
                    segment_customers_display['物料多样性'] = segment_customers_display['物料多样性'].round(0)
                    segment_customers_display['物料销售比率'] = segment_customers_display['物料销售比率'].round(2)

                    st.dataframe(
                        segment_customers_display,
                        use_container_width=True,
                        column_config={
                            "物料总成本": st.column_config.NumberColumn("物料总成本(元)", format="¥%.2f"),
                            "销售总额": st.column_config.NumberColumn("销售总额(元)", format="¥%.2f"),
                            "ROI": st.column_config.NumberColumn("ROI", format="%.2f"),
                            "物料多样性": st.column_config.NumberColumn("物料多样性(种)", format="%d"),
                            "物料销售比率": st.column_config.NumberColumn("物料销售比率(%)", format="%.2f%%")
                        }
                    )

                    # 提供下载各分层客户数据的链接
                    st.markdown(create_download_link(segment_customers_display, f"{segment_name}客户数据"),
                                unsafe_allow_html=True)

            # 分层客户与物料类别偏好匹配分析
            st.markdown('<h3 class="sub-header">客户价值分层与物料类别偏好分析</h3>', unsafe_allow_html=True)

            # 按客户价值分层和物料类别计算平均ROI
            segment_category_roi = segment_material_usage.groupby(['客户价值分层', '物料类别']).agg({
                '物料成本': 'sum',
                '求和项:数量（箱）': 'sum'
            }).reset_index()

            # 匹配销售数据计算ROI
            # 为简化处理，这里假设可以通过客户价值分层和物料类别近似计算ROI
            # 实际应用中可能需要更复杂的匹配逻辑
            segment_category_roi['ROI'] = segment_category_roi.apply(
                lambda row: segment_summary[segment_summary['客户价值分层'] == row['客户价值分层']]['ROI'].values[0],
                axis=1
            )

            # 创建热力图
            segment_heatmap = px.density_heatmap(
                segment_material_usage,
                x='物料类别',
                y='客户价值分层',
                z='物料成本',
                title='客户价值分层与物料类别偏好热力图',
                color_continuous_scale='Viridis',
                histfunc='sum'
            )

            segment_heatmap.update_layout(
                xaxis_title="物料类别",
                yaxis_title="客户价值分层"
            )

            st.plotly_chart(segment_heatmap, use_container_width=True)

            st.markdown(
                '<div class="highlight-box">热力图展示了不同价值分层客户对各类物料的偏好程度。颜色越深表示该分层客户对该类物料的投入越多。通过分析这些偏好模式，可以制定更有针对性的物料投放策略，将合适的物料投放给合适的客户群体。</div>',
                unsafe_allow_html=True)

            # ======= 季节性分析标签页 =======
        with tab6:
            st.markdown(
                '<div class="info-box">本页面分析物料投入和销售产出的季节性变化模式，帮助您识别最佳物料投放时机，实现资源的最优配置。</div>',
                unsafe_allow_html=True)

            # 按月份和季度分析数据
            st.markdown('<h3 class="sub-header">月度与季度分析</h3>', unsafe_allow_html=True)

            # 月度分析 - 所有月份的物料成本和销售额
            monthly_analysis = material_data[
                (material_data['所属区域'].isin(selected_regions)) &
                (material_data['省份'].isin(selected_provinces))
                ].groupby(['月份名', '月份', '月度名称'])['物料成本'].sum().reset_index()

            monthly_sales = sales_data[
                (sales_data['所属区域'].isin(selected_regions)) &
                (sales_data['省份'].isin(selected_provinces))
                ].groupby(['月份名', '月份', '月度名称'])['销售金额'].sum().reset_index()

            monthly_analysis = pd.merge(monthly_analysis, monthly_sales, on=['月份名', '月份', '月度名称'],
                                        how='outer').fillna(0)
            monthly_analysis = monthly_analysis.sort_values('月份')

            # 计算ROI和环比变化
            monthly_analysis['ROI'] = monthly_analysis['销售金额'] / monthly_analysis['物料成本'].replace(0, np.nan)
            monthly_analysis['ROI'].fillna(0, inplace=True)
            monthly_analysis['物料成本环比'] = monthly_analysis['物料成本'].pct_change() * 100
            monthly_analysis['销售金额环比'] = monthly_analysis['销售金额'].pct_change() * 100

            # 格式化数据
            monthly_analysis['物料成本'] = monthly_analysis['物料成本'].round(2)
            monthly_analysis['销售金额'] = monthly_analysis['销售金额'].round(2)
            monthly_analysis['ROI'] = monthly_analysis['ROI'].round(2)
            monthly_analysis['物料成本环比'] = monthly_analysis['物料成本环比'].round(2)
            monthly_analysis['销售金额环比'] = monthly_analysis['销售金额环比'].round(2)

            # 创建月度趋势图
            monthly_trend = make_subplots(specs=[[{"secondary_y": True}]])

            monthly_trend.add_trace(
                go.Bar(
                    x=monthly_analysis['月度名称'],
                    y=monthly_analysis['物料成本'],
                    name='物料成本',
                    marker_color='rgba(58, 71, 80, 0.6)',
                    hovertemplate='月份: %{x}<br>物料成本: %{y:.2f}元<extra></extra>'
                ),
                secondary_y=False
            )

            monthly_trend.add_trace(
                go.Bar(
                    x=monthly_analysis['月度名称'],
                    y=monthly_analysis['销售金额'],
                    name='销售金额',
                    marker_color='rgba(246, 78, 139, 0.6)',
                    hovertemplate='月份: %{x}<br>销售金额: %{y:.2f}元<extra></extra>'
                ),
                secondary_y=False
            )

            monthly_trend.add_trace(
                go.Scatter(
                    x=monthly_analysis['月度名称'],
                    y=monthly_analysis['ROI'],
                    name='ROI',
                    mode='lines+markers+text',
                    line=dict(color='rgb(25, 118, 210)', width=3),
                    marker=dict(size=8),
                    text=monthly_analysis['ROI'].round(2),
                    textposition='top center',
                    hovertemplate='月份: %{x}<br>ROI: %{y:.2f}<extra></extra>'
                ),
                secondary_y=True
            )

            monthly_trend.update_layout(
                title_text='月度物料成本与销售金额趋势',
                barmode='group',
                xaxis_title="月份",
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                )
            )

            monthly_trend.update_yaxes(title_text='金额 (元)', secondary_y=False)
            monthly_trend.update_yaxes(title_text='ROI', secondary_y=True)

            st.plotly_chart(monthly_trend, use_container_width=True)

            # 季度分析
            st.markdown('<h4>季度分析</h4>', unsafe_allow_html=True)

            # 按季度汇总数据
            quarterly_analysis = material_data[
                (material_data['所属区域'].isin(selected_regions)) &
                (material_data['省份'].isin(selected_provinces))
                ].groupby('季度')['物料成本'].sum().reset_index()

            quarterly_sales = sales_data[
                (sales_data['所属区域'].isin(selected_regions)) &
                (sales_data['省份'].isin(selected_provinces))
                ].groupby('季度')['销售金额'].sum().reset_index()

            quarterly_analysis = pd.merge(quarterly_analysis, quarterly_sales, on='季度', how='outer').fillna(0)
            quarterly_analysis = quarterly_analysis.sort_values('季度')

            # 计算ROI
            quarterly_analysis['ROI'] = quarterly_analysis['销售金额'] / quarterly_analysis['物料成本'].replace(0,
                                                                                                                np.nan)
            quarterly_analysis['ROI'].fillna(0, inplace=True)

            # 计算占比
            quarterly_analysis['物料成本占比'] = (
                        quarterly_analysis['物料成本'] / quarterly_analysis['物料成本'].sum() * 100).round(2)
            quarterly_analysis['销售金额占比'] = (
                        quarterly_analysis['销售金额'] / quarterly_analysis['销售金额'].sum() * 100).round(2)

            # 格式化数据
            quarterly_analysis['物料成本'] = quarterly_analysis['物料成本'].round(2)
            quarterly_analysis['销售金额'] = quarterly_analysis['销售金额'].round(2)
            quarterly_analysis['ROI'] = quarterly_analysis['ROI'].round(2)

            # 创建季度对比图
            col1, col2 = st.columns(2)

            with col1:
                # 季度柱状图
                quarterly_bar = px.bar(
                    quarterly_analysis,
                    x='季度',
                    y=['物料成本', '销售金额'],
                    title='季度物料成本与销售金额对比',
                    barmode='group',
                    text_auto='.2f'
                )

                quarterly_bar.update_layout(
                    xaxis_title="季度",
                    yaxis_title="金额 (元)"
                )

                st.plotly_chart(quarterly_bar, use_container_width=True)

            with col2:
                # 季度ROI图
                quarterly_roi = px.line(
                    quarterly_analysis,
                    x='季度',
                    y='ROI',
                    title='季度ROI变化',
                    markers=True,
                    text='ROI'
                )

                quarterly_roi.update_traces(
                    texttemplate='%{text:.2f}',
                    textposition='top center',
                    marker=dict(size=12),
                    line=dict(width=3)
                )

                quarterly_roi.update_layout(
                    xaxis_title="季度",
                    yaxis_title="ROI"
                )

                # 添加参考线 - ROI=1
                quarterly_roi.add_shape(
                    type="line",
                    x0=0.5, y0=1,
                    x1=4.5, y1=1,
                    line=dict(color="red", width=2, dash="dash")
                )

                st.plotly_chart(quarterly_roi, use_container_width=True)

            # 显示季度分析表格
            st.dataframe(
                quarterly_analysis,
                use_container_width=True,
                column_config={
                    "季度": st.column_config.NumberColumn("季度", format="%d"),
                    "物料成本": st.column_config.NumberColumn("物料成本(元)", format="¥%.2f"),
                    "销售金额": st.column_config.NumberColumn("销售金额(元)", format="¥%.2f"),
                    "ROI": st.column_config.NumberColumn("ROI", format="%.2f"),
                    "物料成本占比": st.column_config.NumberColumn("物料成本占比(%)", format="%.2f%%"),
                    "销售金额占比": st.column_config.NumberColumn("销售金额占比(%)", format="%.2f%%")
                }
            )

            st.markdown(
                '<div class="highlight-box">通过月度和季度分析可以发现明显的物料投入和销售产出的季节性模式。通常有1-2个季度表现特别突出，ROI较高，应该在这些季度适当增加物料投入；而在销售淡季，应减少物料投放以控制成本。</div>',
                unsafe_allow_html=True)

            # 物料投放时滞效应分析
            st.markdown('<h3 class="sub-header">物料投放时滞效应分析</h3>', unsafe_allow_html=True)

            # 创建物料投放与销售的交叉关联图
            monthly_corr = monthly_analysis.copy()

            # 计算物料成本滞后1-3个月的销售效果
            for lag in range(1, 4):
                if len(monthly_corr) > lag:
                    monthly_corr[f'物料成本滞后{lag}月'] = monthly_corr['物料成本'].shift(lag)

                    # 计算相关性(避免除以0)
                    if monthly_corr[f'物料成本滞后{lag}月'].std() > 0 and monthly_corr['销售金额'].std() > 0:
                        correlation = monthly_corr['销售金额'].corr(monthly_corr[f'物料成本滞后{lag}月'])
                    else:
                        correlation = 0

                    monthly_corr[f'滞后{lag}月相关性'] = correlation

            # 滞后效应柱状图
            lag_corr = []
            lag_labels = []

            for lag in range(1, 4):
                if f'滞后{lag}月相关性' in monthly_corr.columns:
                    # 取非NaN的平均值
                    corr_value = monthly_corr[f'滞后{lag}月相关性'].dropna().mean()
                    lag_corr.append(corr_value)
                    lag_labels.append(f'滞后{lag}月')

            if lag_corr:  # 确保有数据
                lag_df = pd.DataFrame({
                    '滞后期': lag_labels,
                    '相关性': lag_corr
                })

                lag_chart = px.bar(
                    lag_df,
                    x='滞后期',
                    y='相关性',
                    title='物料投入与销售产出滞后相关性',
                    text='相关性',
                    color='相关性',
                    color_continuous_scale='RdBu',
                    range_color=[-1, 1]
                )

                lag_chart.update_traces(texttemplate='%{text:.4f}', textposition='outside')
                lag_chart.update_layout(
                    xaxis_title="滞后期",
                    yaxis_title="相关系数"
                )

                st.plotly_chart(lag_chart, use_container_width=True)

                # 找出最佳滞后期
                best_lag_idx = np.argmax(np.abs(lag_corr))
                best_lag = lag_labels[best_lag_idx]
                best_corr = lag_corr[best_lag_idx]

                st.markdown(
                    f'<div class="highlight-box">通过分析物料投入与销售产出的时间序列相关性，发现<strong>{best_lag}</strong>的相关系数最高，为<strong>{best_corr:.4f}</strong>。这表明物料投放后大约需要{best_lag.replace("滞后", "").replace("月", "")}个月才能显著影响销售业绩，应据此调整物料投放节奏。</div>',
                    unsafe_allow_html=True)
            else:
                st.warning("数据不足，无法进行滞后效应分析。请确保有足够多的月度数据。")

            # 物料类别的季节性分析
            st.markdown('<h3 class="sub-header">物料类别季节性分析</h3>', unsafe_allow_html=True)

            # 按月份和物料类别分析
            seasonal_category = filtered_material.groupby(['月度名称', '月份', '物料类别'])[
                '物料成本'].sum().reset_index()
            seasonal_category = seasonal_category.sort_values('月份')

            # 创建物料类别季节性热力图
            category_heatmap = px.density_heatmap(
                seasonal_category,
                x='月度名称',
                y='物料类别',
                z='物料成本',
                title='物料类别月度投放热力图',
                color_continuous_scale='Viridis'
            )

            category_heatmap.update_layout(
                xaxis_title="月份",
                yaxis_title="物料类别"
            )

            st.plotly_chart(category_heatmap, use_container_width=True)

            # 按物料类别计算月度分布
            category_monthly_dist = filtered_material.groupby(['物料类别', '月度名称'])['物料成本'].sum().reset_index()
            category_total = filtered_material.groupby('物料类别')['物料成本'].sum().reset_index()
            category_total.rename(columns={'物料成本': '类别总成本'}, inplace=True)

            category_monthly_dist = pd.merge(category_monthly_dist, category_total, on='物料类别', how='left')
            category_monthly_dist['占比'] = (
                        category_monthly_dist['物料成本'] / category_monthly_dist['类别总成本'] * 100).round(2)

            # 创建物料类别月度分布图
            category_monthly_chart = px.line(
                category_monthly_dist,
                x='月度名称',
                y='占比',
                color='物料类别',
                title='各物料类别月度分布',
                markers=True
            )

            category_monthly_chart.update_layout(
                xaxis_title="月份",
                yaxis_title="占总成本百分比(%)"
            )

            st.plotly_chart(category_monthly_chart, use_container_width=True)

            st.markdown(
                '<div class="highlight-box">物料类别季节性分析展示了不同物料类别在各月份的投放情况。热力图中颜色越深表示该月份对应物料类别投入越多。线图展示了各物料类别的月度分布趋势。通过这些分析，可以发现不同物料类别有其独特的季节性模式，应根据这些模式优化物料投放时机。</div>',
                unsafe_allow_html=True)

            # 提供下载季节性分析数据的链接
            st.markdown(create_download_link(monthly_analysis, "月度物料销售分析数据"), unsafe_allow_html=True)
            st.markdown(create_download_link(quarterly_analysis, "季度物料销售分析数据"), unsafe_allow_html=True)

            # ======= 优化建议标签页 =======
        with tab7:
            st.markdown(
                '<div class="info-box">本页面基于数据分析结果，提供物料投放优化建议，帮助您提高物料使用效率和销售业绩。这些建议包括物料投放策略、客户分层管理、季节性调整等多个方面。</div>',
                unsafe_allow_html=True)

            # 1. 物料投放总体优化建议
            st.markdown('<h3 class="sub-header">物料投放总体优化建议</h3>', unsafe_allow_html=True)

            # 计算当前ROI和物料销售比率
            current_roi = filtered_distributor['ROI'].mean() if len(filtered_distributor) > 0 else 0
            current_ratio = filtered_distributor['物料销售比率'].mean() if len(filtered_distributor) > 0 else 0

            # ROI评估
            roi_status = "优秀" if current_roi >= 2.0 else "良好" if current_roi >= 1.0 else "需改进"
            roi_color = "success-metric" if current_roi >= 2.0 else "warning-metric" if current_roi >= 1.0 else "danger-metric"

            # 物料销售比率评估
            ratio_status = "优秀" if current_ratio <= 30 else "良好" if current_ratio <= 50 else "需改进"
            ratio_color = "success-metric" if current_ratio <= 30 else "warning-metric" if current_ratio <= 50 else "danger-metric"

            # 显示总体评估
            st.markdown(f'<h4>当前物料投放整体评估</h4>', unsafe_allow_html=True)
            col1, col2 = st.columns(2)

            with col1:
                st.markdown(
                    f'<div class="metric-card"><h3>投资回报率(ROI): <span class="{roi_color}">{current_roi:.2f}</span></h3><p>状态: {roi_status}</p></div>',
                    unsafe_allow_html=True)

            with col2:
                st.markdown(
                    f'<div class="metric-card"><h3>物料销售比率: <span class="{ratio_color}">{current_ratio:.2f}%</span></h3><p>状态: {ratio_status}</p></div>',
                    unsafe_allow_html=True)

            # 总体优化建议
            overall_recom = [
                "**物料投放总量调整**：根据ROI表现，调整物料总投放量。当ROI<1时，应减少总投放量；当ROI>2时，可适度增加投放以扩大销售。",
                "**物料类别优化**：增加ROI较高物料类别的投放比例，减少ROI较低类别的投放。",
                "**客户差异化投放**：对高价值客户和成长型客户优先配置物料资源，对低效型客户进行物料使用培训后再投放。",
                "**季节性调整**：根据销售季节性，在销售旺季前1-2个月加大物料投放，淡季适当减少投放。",
                "**物料使用培训**：定期为经销商提供物料使用培训，提高物料使用效率。"
            ]

            st.markdown('<div class="highlight-box">', unsafe_allow_html=True)
            for recom in overall_recom:
                st.markdown(recom)
            st.markdown('</div>', unsafe_allow_html=True)

            # 2. 物料类别优化建议
            st.markdown('<h3 class="sub-header">物料类别优化建议</h3>', unsafe_allow_html=True)

            # 获取物料类别ROI数据
            if 'category_roi' in locals():
                # 按ROI排序
                category_sorted = category_roi.sort_values('ROI', ascending=False).copy()

                # 增加、减少和优化的类别
                increase_categories = category_sorted[category_sorted['ROI'] >= 2.0]['物料类别'].tolist()
                maintain_categories = category_sorted[(category_sorted['ROI'] >= 1.0) & (category_sorted['ROI'] < 2.0)][
                    '物料类别'].tolist()
                reduce_categories = category_sorted[category_sorted['ROI'] < 1.0]['物料类别'].tolist()

                # 创建优化建议
                st.markdown('<h4>物料类别投放调整建议</h4>', unsafe_allow_html=True)

                col1, col2, col3 = st.columns(3)

                with col1:
                    st.markdown('<div style="border-left: 4px solid #4CAF50; padding-left: 10px;">',
                                unsafe_allow_html=True)
                    st.markdown('#### 增加投放')
                    if increase_categories:
                        for cat in increase_categories:
                            roi_val = category_sorted[category_sorted['物料类别'] == cat]['ROI'].values[0]
                            st.markdown(f"**{cat}** (ROI: {roi_val:.2f})")
                        st.markdown("这些物料类别ROI≥2.0，投入产出效果优秀，建议增加投放比例。")
                    else:
                        st.markdown("暂无ROI≥2.0的物料类别，建议优化现有物料使用方式。")
                    st.markdown('</div>', unsafe_allow_html=True)

                with col2:
                    st.markdown('<div style="border-left: 4px solid #FFC107; padding-left: 10px;">',
                                unsafe_allow_html=True)
                    st.markdown('#### 维持投放')
                    if maintain_categories:
                        for cat in maintain_categories:
                            roi_val = category_sorted[category_sorted['物料类别'] == cat]['ROI'].values[0]
                            st.markdown(f"**{cat}** (ROI: {roi_val:.2f})")
                        st.markdown("这些物料类别1.0≤ROI<2.0，投入产出效果良好，建议维持投放并优化使用方式。")
                    else:
                        st.markdown("暂无1.0≤ROI<2.0的物料类别。")
                    st.markdown('</div>', unsafe_allow_html=True)

                with col3:
                    st.markdown('<div style="border-left: 4px solid #F44336; padding-left: 10px;">',
                                unsafe_allow_html=True)
                    st.markdown('#### 减少投放')
                    if reduce_categories:
                        for cat in reduce_categories:
                            roi_val = category_sorted[category_sorted['物料类别'] == cat]['ROI'].values[0]
                            st.markdown(f"**{cat}** (ROI: {roi_val:.2f})")
                        st.markdown("这些物料类别ROI<1.0，投入产出效果不佳，建议减少投放并分析原因。")
                    else:
                        st.markdown("暂无ROI<1.0的物料类别，现有物料类别投放效果良好。")
                    st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.warning("物料类别ROI数据不足，无法生成具体的物料类别优化建议。")

            # 3. 客户分层管理建议
            st.markdown('<h3 class="sub-header">客户分层管理建议</h3>', unsafe_allow_html=True)

            customer_strategies = {
                '高价值客户': {
                    '特点': '物料使用效率高，销售表现优异',
                    '策略': [
                        "优先分配优质物料资源",
                        "提供个性化物料组合方案",
                        "定期回访，收集最佳实践",
                        "建立示范案例，推广成功经验",
                        "提供专属物料使用培训服务"
                    ]
                },
                '成长型客户': {
                    '特点': '物料使用有效，销售潜力大',
                    '策略': [
                        "适度增加物料投放量",
                        "优化物料组合结构",
                        "加强物料使用培训",
                        "提供销售技巧指导",
                        "定期评估成长情况，适时调整支持力度"
                    ]
                },
                '稳定型客户': {
                    '特点': '物料投入产出平衡，销售表现稳定',
                    '策略': [
                        "保持稳定的物料供应",
                        "适当提升物料多样性",
                        "分享高效使用物料的案例",
                        "引导尝试ROI较高的物料类别",
                        "定期检视物料使用效率"
                    ]
                },
                '低效型客户': {
                    '特点': '物料投入未产生有效回报',
                    '策略': [
                        "减少物料总投放量",
                        "重点提供基础物料",
                        "开展物料使用专项培训",
                        "分析低效原因，提供针对性指导",
                        "实行物料与销售挂钩机制"
                    ]
                }
            }

            # 使用卡片形式展示客户分层管理建议
            col1, col2 = st.columns(2)

            with col1:
                for segment, strategy in list(customer_strategies.items())[:2]:
                    st.markdown(
                        f'<div style="border: 1px solid #ddd; border-radius: 5px; padding: 15px; margin-bottom: 15px;">',
                        unsafe_allow_html=True)
                    st.markdown(f'<h4>{segment}</h4>', unsafe_allow_html=True)
                    st.markdown(f'<p><strong>特点：</strong>{strategy["特点"]}</p>', unsafe_allow_html=True)
                    st.markdown('<p><strong>管理策略：</strong></p>', unsafe_allow_html=True)
                    for item in strategy['策略']:
                        st.markdown(f"- {item}")
                    st.markdown('</div>', unsafe_allow_html=True)

            with col2:
                for segment, strategy in list(customer_strategies.items())[2:]:
                    st.markdown(
                        f'<div style="border: 1px solid #ddd; border-radius: 5px; padding: 15px; margin-bottom: 15px;">',
                        unsafe_allow_html=True)
                    st.markdown(f'<h4>{segment}</h4>', unsafe_allow_html=True)
                    st.markdown(f'<p><strong>特点：</strong>{strategy["特点"]}</p>', unsafe_allow_html=True)
                    st.markdown('<p><strong>管理策略：</strong></p>', unsafe_allow_html=True)
                    for item in strategy['策略']:
                        st.markdown(f"- {item}")
                    st.markdown('</div>', unsafe_allow_html=True)

            # 4. 季节性调整建议
            st.markdown('<h3 class="sub-header">季节性调整建议</h3>', unsafe_allow_html=True)

            # 季度投放建议
            if 'quarterly_analysis' in locals():
                # 按ROI识别最佳和最差季度
                best_quarter = quarterly_analysis.loc[quarterly_analysis['ROI'].idxmax()]
                worst_quarter = quarterly_analysis.loc[quarterly_analysis['ROI'].idxmin()]

                st.markdown('<h4>季度物料投放策略</h4>', unsafe_allow_html=True)

                col1, col2 = st.columns(2)

                with col1:
                    st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                    st.markdown(f'<h4>最佳季度: Q{best_quarter["季度"]}</h4>', unsafe_allow_html=True)
                    st.markdown(f'<p>ROI: <span class="success-metric">{best_quarter["ROI"]:.2f}</span></p>',
                                unsafe_allow_html=True)
                    st.markdown(f'<p>物料成本: ¥{best_quarter["物料成本"]:.2f}</p>', unsafe_allow_html=True)
                    st.markdown(f'<p>销售金额: ¥{best_quarter["销售金额"]:.2f}</p>', unsafe_allow_html=True)
                    st.markdown('<p><strong>建议：</strong></p>', unsafe_allow_html=True)
                    st.markdown("- 增加物料投放总量15-20%")
                    st.markdown("- 优先安排高ROI物料类别")
                    st.markdown("- 提前1-2个月开始准备物料")
                    st.markdown("- 重点覆盖高价值客户和成长型客户")
                    st.markdown('</div>', unsafe_allow_html=True)

                with col2:
                    st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                    st.markdown(f'<h4>最差季度: Q{worst_quarter["季度"]}</h4>', unsafe_allow_html=True)
                    st.markdown(f'<p>ROI: <span class="danger-metric">{worst_quarter["ROI"]:.2f}</span></p>',
                                unsafe_allow_html=True)
                    st.markdown(f'<p>物料成本: ¥{worst_quarter["物料成本"]:.2f}</p>', unsafe_allow_html=True)
                    st.markdown(f'<p>销售金额: ¥{worst_quarter["销售金额"]:.2f}</p>', unsafe_allow_html=True)
                    st.markdown('<p><strong>建议：</strong></p>', unsafe_allow_html=True)
                    st.markdown("- 减少物料投放总量20-30%")
                    st.markdown("- 只保留必要的基础物料")
                    st.markdown("- 清理库存，避免物料积压")
                    st.markdown("- 开展物料使用培训，为下一个旺季做准备")
                    st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.warning("季度分析数据不足，无法生成季节性调整建议。")

            # 月度投放节奏建议
            st.markdown('<h4>月度物料投放节奏建议</h4>', unsafe_allow_html=True)

            monthly_strategy = [
                "**提前规划**：基于历史销售数据，提前2-3个月规划物料需求，确保及时到位。",
                "**梯次投放**：将物料分批次投放，旺季前逐步增加，避免一次性投入过多导致使用效率低。",
                "**旺季前倾**：在销售旺季前1-2个月加大物料投放力度，确保销售人员有足够时间熟悉和使用物料。",
                "**淡季控量**：销售淡季减少物料投放50%以上，只保留必要的基础物料。",
                "**滞后期调整**：根据分析发现的物料投放滞后效应，提前调整物料投放时机，使销售效果最大化。"
            ]

            st.markdown('<div class="highlight-box">', unsafe_allow_html=True)
            for strategy in monthly_strategy:
                st.markdown(strategy)
            st.markdown('</div>', unsafe_allow_html=True)

            # 5. 重点关注的经销商
            st.markdown('<h3 class="sub-header">重点关注的经销商</h3>', unsafe_allow_html=True)

            # 筛选需要重点关注的经销商
            if len(filtered_distributor) > 0:
                # 高销售额但低ROI的经销商（潜力客户）
                potential_customers = filtered_distributor[
                    (filtered_distributor['销售总额'] > filtered_distributor['销售总额'].quantile(0.75)) &
                    (filtered_distributor['ROI'] < 1.0)
                    ].sort_values('销售总额', ascending=False).head(5)

                # 高ROI但低销售额的经销商（成长客户）
                growth_customers = filtered_distributor[
                    (filtered_distributor['ROI'] > 2.0) &
                    (filtered_distributor['销售总额'] < filtered_distributor['销售总额'].median())
                    ].sort_values('ROI', ascending=False).head(5)

                # 双高客户（标杆客户）
                benchmark_customers = filtered_distributor[
                    (filtered_distributor['ROI'] > 2.0) &
                    (filtered_distributor['销售总额'] > filtered_distributor['销售总额'].quantile(0.75))
                    ].sort_values('ROI', ascending=False).head(5)

                # 双低客户（问题客户）
                problem_customers = filtered_distributor[
                    (filtered_distributor['ROI'] < 0.5) &
                    (filtered_distributor['物料总成本'] > filtered_distributor['物料总成本'].quantile(0.5))
                    ].sort_values('物料总成本', ascending=False).head(5)

                # 创建重点关注客户表格
                focus_customers = pd.DataFrame({
                    '客户类型': ['潜力客户', '成长客户', '标杆客户', '问题客户'],
                    '描述': [
                        '销售额高但ROI低，物料使用效率有待提高',
                        'ROI高但销售额低，具有成长潜力',
                        'ROI高且销售额高，可作为最佳实践标杆',
                        'ROI极低且物料投入较多，亟需干预'
                    ],
                    '客户名单': [
                        ', '.join(potential_customers['经销商名称'].tolist()) if len(potential_customers) > 0 else '无',
                        ', '.join(growth_customers['经销商名称'].tolist()) if len(growth_customers) > 0 else '无',
                        ', '.join(benchmark_customers['经销商名称'].tolist()) if len(benchmark_customers) > 0 else '无',
                        ', '.join(problem_customers['经销商名称'].tolist()) if len(problem_customers) > 0 else '无'
                    ],
                    '干预策略': [
                        '诊断物料使用问题，提供专项培训，调整物料结构',
                        '增加物料投放量，扩大销售规模，挖掘增长潜力',
                        '研究最佳实践，推广成功经验，作为标杆案例',
                        '大幅减少物料投放，培训后再投放，定期评估改进情况'
                    ]
                })

                st.dataframe(focus_customers, use_container_width=True)
            else:
                st.warning("经销商数据不足，无法生成重点关注的经销商建议。")

            # 6. 执行计划
            st.markdown('<h3 class="sub-header">物料优化执行计划</h3>', unsafe_allow_html=True)

            execution_plan = [
                {
                    '阶段': '第一阶段：诊断评估（1个月）',
                    '关键任务': [
                        "对所有经销商进行物料使用效率评估",
                        "识别高效和低效经销商，分析差异原因",
                        "评估各物料类别ROI，识别高效和低效物料",
                        "制定物料投放优化目标"
                    ]
                },
                {
                    '阶段': '第二阶段：策略制定（1个月）',
                    '关键任务': [
                        "按客户价值分层制定差异化物料投放策略",
                        "优化物料类别组合结构",
                        "设计季节性物料投放计划",
                        "制定物料使用培训方案"
                    ]
                },
                {
                    '阶段': '第三阶段：试点实施（2个月）',
                    '关键任务': [
                        "选择20%的经销商进行试点",
                        "实施优化后的物料投放方案",
                        "开展物料使用培训",
                        "建立物料使用跟踪机制"
                    ]
                },
                {
                    '阶段': '第四阶段：评估调整（1个月）',
                    '关键任务': [
                        "评估试点效果",
                        "总结经验教训",
                        "优化调整方案",
                        "准备全面推广"
                    ]
                },
                {
                    '阶段': '第五阶段：全面推广（3个月）',
                    '关键任务': [
                        "向所有经销商推广优化方案",
                        "分批实施培训计划",
                        "建立定期评估机制",
                        "完善物料管理系统"
                    ]
                }
            ]

            # 创建执行计划展示
            for phase in execution_plan:
                st.markdown(
                    f'<div style="border: 1px solid #ddd; border-radius: 5px; padding: 15px; margin-bottom: 15px;">',
                    unsafe_allow_html=True)
                st.markdown(f'<h4>{phase["阶段"]}</h4>', unsafe_allow_html=True)
                st.markdown('<p><strong>关键任务：</strong></p>', unsafe_allow_html=True)
                for task in phase['关键任务']:
                    st.markdown(f"- {task}")
                st.markdown('</div>', unsafe_allow_html=True)

            # 7. 预期效果
            st.markdown('<h3 class="sub-header">预期效果</h3>', unsafe_allow_html=True)

            # 创建预期效果图表
            expected_results = pd.DataFrame({
                '指标': ['物料总成本', '销售总额', 'ROI', '高价值客户占比', '物料销售比率'],
                '当前值': [
                    filtered_material['物料成本'].sum() if len(filtered_material) > 0 else 0,
                    filtered_sales['销售金额'].sum() if len(filtered_sales) > 0 else 0,
                    current_roi,
                    (filtered_distributor['客户价值分层'] == '高价值客户').mean() * 100 if len(
                        filtered_distributor) > 0 else 0,
                    current_ratio
                ],
                '目标值': [
                    filtered_material['物料成本'].sum() * 0.9 if len(filtered_material) > 0 else 0,  # 降低10%
                    filtered_sales['销售金额'].sum() * 1.2 if len(filtered_sales) > 0 else 0,  # 提升20%
                    current_roi * 1.5 if current_roi > 0 else 2.0,  # 提升50%或设为2.0
                    min(((filtered_distributor['客户价值分层'] == '高价值客户').mean() * 100 * 1.5), 30) if len(
                        filtered_distributor) > 0 else 25,  # 提升50%或最高30%
                    max(current_ratio * 0.7, 20) if current_ratio > 0 else 20  # 降低30%或最低20%
                ]
            })

            # 格式化数据
            expected_results['当前值'] = expected_results.apply(
                lambda row: f"¥{row['当前值']:.2f}" if row['指标'] in ['物料总成本', '销售总额'] else
                f"{row['当前值']:.2f}" if row['指标'] == 'ROI' else
                f"{row['当前值']:.2f}%" if row['指标'] in ['高价值客户占比', '物料销售比率'] else
                row['当前值'],
                axis=1
            )

            expected_results['目标值'] = expected_results.apply(
                lambda row: f"¥{row['目标值']:.2f}" if row['指标'] in ['物料总成本', '销售总额'] else
                f"{row['目标值']:.2f}" if row['指标'] == 'ROI' else
                f"{row['目标值']:.2f}%" if row['指标'] in ['高价值客户占比', '物料销售比率'] else
                row['目标值'],
                axis=1
            )

            # 显示预期效果表格
            st.dataframe(expected_results, use_container_width=True)

            # 最终总结
            st.markdown('<div class="highlight-box">', unsafe_allow_html=True)
            st.markdown("""
                        通过实施以上优化建议，预计能够显著提升物料投放效率，降低物料总成本，同时提升销售业绩。核心目标是通过优化物料投放结构、客户分层管理和季节性调整，将ROI提升50%以上，实现"用更少的物料创造更多的销售"。

                        物料优化是一个持续过程，建议每季度进行一次评估，根据实际效果不断调整优化策略。同时，收集经销商反馈，了解一线需求，使物料投放更加贴合市场实际情况。
                        """)
            st.markdown('</div>', unsafe_allow_html=True)

            # 页脚显示数据更新时间
        st.markdown("---")
        st.caption(f"数据最后更新时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")

if __name__ == '__main__':
        main()