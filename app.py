import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO

st.set_page_config(
    page_title="校园生活备忘录",
    page_icon="🎒",
    layout="wide"
)

st.title("🎒 我的校园生活备忘录")
st.caption("快递取件码、待办、宿舍报修，一个网页全记下来，不用翻微信聊天记录")

# 初始化数据
if 'express_list' not in st.session_state:
    st.session_state.express_list = []
if 'todo_list' not in st.session_state:
    st.session_state.todo_list = []
if 'repair_list' not in st.session_state:
    st.session_state.repair_list = []

menu = st.sidebar.radio("功能导航", ["📦 快递取件码", "✅ 待办事项", "🔧 宿舍报修", "💾 数据备份"])

# ------------------- 快递 -------------------
if menu == "📦 快递取件码":
    st.header("📦 快递取件码记录")
    with st.form("add_express", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            company = st.selectbox("快递公司",
                                   ["菜鸟驿站", "京东快递", "顺丰", "圆通", "中通", "韵达", "极兔", "邮政", "其他"])
            code = st.text_input("取件码", placeholder="例如：1-3-2056")
        with c2:
            addr = st.text_input("取件地点", placeholder="例如：南区菜鸟驿站")
            note = st.text_input("备注（可选）", placeholder="买了什么东西")
        submitted = st.form_submit_button("➕ 添加快递")
        if submitted and code:
            st.session_state.express_list.append({
                "快递公司": company,
                "取件码": code,
                "取件地点": addr,
                "备注": note,
                "存入时间": datetime.now().strftime("%Y-%m-%d %H:%M")
            })
            st.success("快递已记录！")

    if st.session_state.express_list:
        df = pd.DataFrame(st.session_state.express_list)
        st.dataframe(df, use_container_width=True, hide_index=True)
        if st.button("🗑️ 清空快递列表"):
            st.session_state.express_list = []
            st.rerun()
    else:
        st.info("还没有快递记录，收到快递把取件码记在这里，不会忘！")

# ------------------- 待办（新增勾选完成） -------------------
elif menu == "✅ 待办事项":
    st.header("✅ 今日待办")
    with st.form("add_todo", clear_on_submit=True):
        task = st.text_input("要做什么事", placeholder="例如：明天交人力资源作业、取快递")
        ddl = st.date_input("截止日期", value=date.today())
        priority = st.select_slider("优先级", ["低", "中", "高"])
        submitted = st.form_submit_button("➕ 添加待办")
        if submitted and task:
            st.session_state.todo_list.append({
                "事项": task,
                "截止日期": ddl.strftime("%Y-%m-%d"),
                "优先级": priority,
                "完成状态": False
            })
            st.success("待办已添加！")

    if st.session_state.todo_list:
        new_todo = []
        for item in st.session_state.todo_list:
            done = st.checkbox(f'【{item["优先级"]} | {item["截止日期"]}】 {item["事项"]}', value=item["完成状态"])
            item["完成状态"] = done
            new_todo.append(item)
        st.session_state.todo_list = new_todo

        if st.button("🗑️ 删除全部已完成"):
            st.session_state.todo_list = [i for i in st.session_state.todo_list if not i["完成状态"]]
            st.rerun()
    else:
        st.info("还没有待办事项，把要做的事记在这里，避免忘事！")

# ------------------- 报修 -------------------
elif menu == "🔧 宿舍报修":
    st.header("🔧 宿舍报修记录")
    with st.form("add_repair", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            item = st.selectbox("报修项目",
                                ["灯坏了", "空调不制冷", "水龙头漏水", "门锁坏了", "网络断了", "插座坏了", "其他"])
            dorm = st.text_input("宿舍号", placeholder="例如：12栋301")
        with c2:
            status = st.selectbox("处理状态", ["已提交", "维修中", "已修好"])
            desc = st.text_area("问题描述", placeholder="简单说一下坏的情况")
        submitted = st.form_submit_button("➕ 登记报修")
        if submitted:
            st.session_state.repair_list.append({
                "报修项目": item,
                "宿舍号": dorm,
                "处理状态": status,
                "问题描述": desc,
                "提交时间": datetime.now().strftime("%Y-%m-%d %H:%M")
            })
            st.success("报修记录已保存！")

    if st.session_state.repair_list:
        df = pd.DataFrame(st.session_state.repair_list)
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("宿舍东西坏了，把报修进度记在这里，不用反复问宿管！")

# ------------------- 备份（补全导入恢复） -------------------
from io import BytesIO
tab1, tab2 = st.tabs(["导出备份", "导入恢复"])
with tab1:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(st.session_state.express_list).to_excel(writer, sheet_name="快递", index=False)
        pd.DataFrame(st.session_state.todo_list).to_excel(writer, sheet_name="待办", index=False)
        pd.DataFrame(st.session_state.repair_list).to_excel(writer, sheet_name="报修", index=False)
    buf.seek(0)
    st.download_button(
        label="📥 下载Excel备份文件",
        data=buf,
        file_name="campus_backup.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
with tab2:
    uploaded_xlsx = st.file_uploader("上传 campus_backup.xlsx", type="xlsx")
    if uploaded_xlsx is not None:
        st.session_state.express_list = pd.read_excel(uploaded_xlsx, sheet_name="快递").to_dict("records")
        st.session_state.todo_list = pd.read_excel(uploaded_xlsx, sheet_name="待办").to_dict("records")
        st.session_state.repair_list = pd.read_excel(uploaded_xlsx, sheet_name="报修").to_dict("records")
        st.success("✅ Excel导入成功！")