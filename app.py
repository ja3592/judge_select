'''
版本号：V1
功能：用于辅助筛选样板车评委，提高效率，具有省心、不纠结的特点。
开发时间：2025.04
'''

import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="样板车评委安排系统", layout="wide")

# 设置页面标题
st.title("样板车评委安排系统")

# 初始化会话状态变量
if 'df' not in st.session_state:
    st.session_state.df = None
if 'selected_company' not in st.session_state:
    st.session_state.selected_company = None
if 'candidate_judges' not in st.session_state:
    st.session_state.candidate_judges = None
if 'final_judges' not in st.session_state:
    st.session_state.final_judges = None
if 'company_list' not in st.session_state:
    st.session_state.company_list = []

# 上传评委信息总表
uploaded_file = st.file_uploader("上传评委信息总表", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.session_state.df = df
        
        # 提取公司列表（假设公司列从第4列开始）
        company_columns = df.columns[3:-1]  # 排除最后一列（已参加评审次数和）
        st.session_state.company_list = company_columns.tolist()
        
        st.success("评委信息总表上传成功！")
        
        # 显示原始数据表
        with st.expander("查看原始评委信息表"):
            st.dataframe(df)
    except Exception as e:
        st.error(f"文件读取错误: {e}")

# 如果数据已上传，显示筛选和操作界面
if st.session_state.df is not None:
    st.markdown("---")
    st.subheader("评委筛选设置")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # 选择待评审公司
        selected_company = st.selectbox(
            "选择待评审公司", 
            options=st.session_state.company_list
        )
        
        # 检查该公司是否已经参加过评审
        if selected_company:
            df = st.session_state.df
            if (df[selected_company] == 1).any() or (df[selected_company] == '1').any():
                st.warning(f"⚠️ {selected_company}已经参加过评审，请重新选择。")
            else:
                st.session_state.selected_company = selected_company
        
        # 设置参评次数上限
        max_reviews = st.number_input("参评次数上限", min_value=0, value=3)
        
        # 设置每家公司最多评委数
        max_judges_per_company = st.number_input("每家公司最多评委数", min_value=1, max_value=5, value=1)
    
    with col2:
        # 排除公司
        excluded_companies = st.multiselect(
            "排除公司（评委来源）", 
            options=st.session_state.df["评委所属公司"].unique()
        )
        
        # 排除评委
        excluded_judges = st.multiselect(
            "排除评委", 
            options=st.session_state.df["评委姓名"].unique()
        )
    
    with col3:
        # 新增：设置组长和组员数量
        num_leaders = st.number_input("组长数量", min_value=1, max_value=5, value=2)
        num_members = st.number_input("组员数量", min_value=1, max_value=10, value=7)
    
    # 生成备选评委函数
    def generate_candidate_judges(current_judges=None, replace_judges_names=None):
        df = st.session_state.df
        
        # 筛选条件
        filtered_df = df[
            (df[selected_company] != 1) &  # 未参评该公司
            (df["已参加评审次数和"] < max_reviews) &  # 参评次数未超过上限
            (~df["评委所属公司"].isin([selected_company])) &  # 不是待评审公司的评委
            (~df["评委所属公司"].isin(excluded_companies)) &  # 不是排除公司的评委
            (~df["评委姓名"].isin(excluded_judges))  # 不是排除的评委
        ]
        
        # 初始化结果和公司计数
        result_judges = []
        company_count = {}
        
        # 如果有当前评委列表，先处理需要保留的评委
        if current_judges is not None and not current_judges.empty:
            # 找出需要保留的评委（未被勾选的评委）
            if replace_judges_names:
                keep_judges = current_judges[~current_judges["评委姓名"].isin(replace_judges_names)]
            else:
                keep_judges = current_judges
            
            # 将保留的评委添加到结果中
            for _, judge in keep_judges.iterrows():
                result_judges.append(judge)
                company = judge["评委所属公司"]
                if company not in company_count:
                    company_count[company] = 0
                company_count[company] += 1
            
            # 从筛选结果中移除当前筛选出来的评委，如果知识移除已保留的，那么去除掉的有可能会被重新选中
            filtered_df = filtered_df[~filtered_df["评委姓名"].isin(current_judges["评委姓名"])]
        
        # 分离组长和组员
        all_leaders = filtered_df[filtered_df["评委类别"] == "组长"]
        all_members = filtered_df[filtered_df["评委类别"] == "组员"]
        
        # 计算已有的组长和组员数量
        current_leaders = [j for j in result_judges if j["评委类别"] == "组长"]
        current_members = [j for j in result_judges if j["评委类别"] == "组员"]
        
        # 随机选择组长，直到达到设定数量
        if len(current_leaders) < num_leaders:
            # 随机打乱顺序
            leaders = all_leaders.sample(frac=1).reset_index(drop=True)
            
            for _, leader in leaders.iterrows():
                company = leader["评委所属公司"]
                if company not in company_count:
                    company_count[company] = 0
                
                if company_count[company] < max_judges_per_company and len(current_leaders) < num_leaders:
                    result_judges.append(leader)
                    current_leaders.append(leader)
                    company_count[company] += 1
                
                if len(current_leaders) >= num_leaders:
                    break
        
        # 随机选择组员，直到达到设定数量
        if len(current_members) < num_members:
            # 随机打乱顺序
            members = all_members.sample(frac=1).reset_index(drop=True)
            
            for _, member in members.iterrows():
                company = member["评委所属公司"]
                if company not in company_count:
                    company_count[company] = 0
                
                if company_count[company] < max_judges_per_company and len(current_members) < num_members:
                    result_judges.append(member)
                    current_members.append(member)
                    company_count[company] += 1
                
                if len(current_members) >= num_members:
                    break
        
        # 合并选中的评委
        if result_judges:
            return pd.DataFrame(result_judges)
        else:
            return pd.DataFrame()
    
    # 生成备选评委按钮
    if st.button("生成备选评委"):
        # 检查所选公司是否已经参加过评审
        df = st.session_state.df
        if (df[selected_company] == 1).any() or (df[selected_company] == '1').any():
            st.error(f"{selected_company}已经参加过评审，请重新选择。")
        else:
            candidate_judges = generate_candidate_judges()
            
            total_required = num_leaders + num_members
            if len(candidate_judges) < total_required:
                st.warning(f"符合条件的评委不足{total_required}名，仅找到{len(candidate_judges)}名评委。")
            
            # 保存备选评委
            st.session_state.candidate_judges = candidate_judges
    
    # 显示备选评委
    if st.session_state.candidate_judges is not None and len(st.session_state.candidate_judges) > 0:
        st.markdown("---")
        st.subheader(f"{selected_company}评委人员备选名单")
        
        # 显示备选评委表格
        candidate_display = st.session_state.candidate_judges[["评委姓名", "评委所属公司", "评委类别", "已参加评审次数和"]].copy()
        
        # 添加选择列
        candidate_display["重新选择"] = False
        
        # 使用可编辑的数据表格
        edited_df = st.data_editor(
            candidate_display,
            column_config={
                "重新选择": st.column_config.CheckboxColumn(
                    "重新选择",
                    help="勾选此项将替换该评委",
                    default=False,
                )
            },
            hide_index=True,
            key="candidate_editor"
        )
        
        # 重新选择按钮
        col1, col2, col3 = st.columns([1, 1, 1])
        
        with col1:
            if st.button("重新选择勾选的评委"):
                # 获取勾选的评委（需要替换的评委）
                replace_judges = edited_df[edited_df["重新选择"] == True]
                
                if not replace_judges.empty:
                    replace_judges_names = replace_judges["评委姓名"].tolist()
                    
                    # 重新生成备选评委，替换勾选的评委
                    new_candidate_judges = generate_candidate_judges(
                        current_judges=st.session_state.candidate_judges,
                        replace_judges_names=replace_judges_names
                    )
                    
                    st.session_state.candidate_judges = new_candidate_judges
                    st.rerun()
                else:
                    st.info("请勾选需要替换的评委")
        
        with col2:
            # 下载备选评委名单 - 直接下载
            output = io.BytesIO()
            download_df = candidate_display.drop(columns=["重新选择"]) # 去掉重新选择列
            download_df.to_excel(output, index=False, sheet_name=f"{selected_company}评委人员备选名单")
            output.seek(0)
            
            st.download_button(
                label="下载备选评委名单",
                data=output,
                file_name=f"{selected_company}评委人员备选名单.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col3:
            # 生成正式评委名单
            if st.button("生成正式评委名单"):
                # 获取选中的评委（勾选"重新选择"的评委）
                selected_rows = edited_df[edited_df["重新选择"] == True]
                
                # 检查是否选择了至少1名组长和4名组员
                selected_leaders = selected_rows[selected_rows["评委类别"] == "组长"]
                selected_members = selected_rows[selected_rows["评委类别"] == "组员"]
                
                if len(selected_leaders) < 1:
                    st.error("正式评委中需要至少1名组长！请增加勾选的组长数量。")
                elif len(selected_members) < 4:
                    st.error("正式评委中需要至少4名组员！请增加勾选的组员数量。")
                else:
                    # 保存正式评委名单
                    final_judges = selected_rows.drop(columns=["重新选择"])
                    st.session_state.final_judges = final_judges
                    
                    st.success("正式评委名单已生成！")
        
        # 显示和下载正式评委名单
        if st.session_state.final_judges is not None:
            st.markdown("---")
            st.subheader(f"{selected_company}评委人员正式名单")
            
            st.dataframe(st.session_state.final_judges)
            
            # 下载正式评委名单 - 直接下载
            output = io.BytesIO()
            st.session_state.final_judges.to_excel(output, index=False, sheet_name=f"{selected_company}评委人员正式名单")
            output.seek(0)
            
            st.download_button(
                label="下载正式评委名单",
                data=output,
                file_name=f"{selected_company}评委人员正式名单.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # 更新评委信息总表
            if st.button("更新评委信息总表"):
                # 获取原始数据
                df = st.session_state.df.copy()
                
                # 更新参评信息
                for _, judge in st.session_state.final_judges.iterrows():
                    judge_name = judge["评委姓名"]
                    
                    # 找到对应评委的行索引
                    judge_idx = df[df["评委姓名"] == judge_name].index[0]
                    
                    # 更新参评公司标记为1
                    df.loc[judge_idx, selected_company] = 1
                    
                    # 重新计算参评次数和
                    company_columns = st.session_state.company_list
                    df.loc[judge_idx, "已参加评审次数和"] = df.loc[judge_idx, company_columns].sum()
                
                # 保存更新后的数据
                st.session_state.df = df
                
                # 下载更新后的评委信息总表 - 直接下载
                output = io.BytesIO()
                current_time = datetime.now().strftime("%Y%m%d%H%M%S")
                df.to_excel(output, index=False, sheet_name="评委信息总表")
                output.seek(0)
                
                st.download_button(
                    label="下载更新后的评委信息总表",
                    data=output,
                    file_name=f"评委信息总表_{current_time}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("评委信息总表已更新！")


# 添加使用帮助
with st.expander("使用帮助"):
    st.markdown("""
    ### 使用步骤
    1. **上传文件**：上传包含评委信息的Excel表格，注意必须采用固定的表头、模版。
    2. **设置参数**：选择待评审公司，设置筛选条件，比如排除的公司、排除的评委、允许评委参评次数等。
    3. **点击生成**：生成备选评委名单后，可以直接下载，也可以继续操作，包括评委的局部调整、确认后评委的下载以及原评委信息表的更新。
    """)
# 添加页脚
st.markdown("---")
st.markdown("<div style='text-align: center;'>© 2025 样板车评委安排系统 | 集团公司样板车评委管理</div>", unsafe_allow_html=True)