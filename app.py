"""
妍發科技 製程管理系統 — Streamlit 主應用
啟動指令: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import os
import sys

# 確保可以 import 同目錄模組
sys.path.insert(0, os.path.dirname(__file__))

import database as db

# ── 頁面設定 ──
st.set_page_config(
    page_title="妍發科技 製程管理系統",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 初始化資料庫 ──
db.init_db()


# ── 側邊欄導航 ──
st.sidebar.title("🏭 妍發科技")
st.sidebar.markdown("**製程管理系統**")
st.sidebar.divider()

page = st.sidebar.radio(
    "功能選單",
    ["📊 儀表板", "📋 訂單管理", "🔧 製程追蹤", "🚚 出貨排程", "🏗️ 廠商統計", "📥 資料匯入"],
    label_visibility="collapsed",
)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 📊 儀表板
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if page == "📊 儀表板":
    st.title("📊 儀表板總覽")

    stats = db.get_dashboard_stats()

    # 上方指標卡片
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("訂單總數", stats['total_orders'])
    col2.metric("製程中", stats['in_progress'], delta=None)
    col3.metric("已出貨", stats['shipped'])
    col4.metric("延遲", stats['delayed'], delta=f"-{stats['delayed']}" if stats['delayed'] > 0 else None, delta_color="inverse")
    col5.metric("暫停/待料", stats['on_hold'])

    st.divider()

    # 金額統計
    col_a, col_b, col_c = st.columns(3)
    col_a.metric("訂單總金額", f"${stats['total_amount']:,.0f}")
    col_b.metric("已出貨金額", f"${stats['shipped_amount']:,.0f}")
    remaining = stats['total_amount'] - stats['shipped_amount']
    col_c.metric("未出貨金額", f"${remaining:,.0f}")

    st.divider()

    # 訂單狀態分布圖
    left, right = st.columns(2)

    with left:
        st.subheader("訂單狀態分布")
        orders = db.get_orders()
        if orders:
            df_orders = pd.DataFrame([dict(r) for r in orders])
            status_counts = df_orders['status'].value_counts()
            st.bar_chart(status_counts)
        else:
            st.info("尚無訂單資料，請先到「資料匯入」匯入 XLS 檔案。")

    with right:
        st.subheader("廠商負載排行")
        vendor_load = db.get_vendor_load()
        if vendor_load:
            df_vendor = pd.DataFrame([dict(r) for r in vendor_load])
            df_vendor = df_vendor.rename(columns={
                'vendor_name': '廠商',
                'total_steps': '總製程數',
                'completed': '已完成',
                'in_progress': '進行中',
                'delayed': '延遲'
            })
            st.dataframe(df_vendor, hide_index=True, use_container_width=True)
        else:
            st.info("尚無廠商資料。")

    # 月出貨趨勢
    st.subheader("月出貨金額趨勢")
    monthly = db.get_monthly_shipment_summary()
    if monthly:
        df_monthly = pd.DataFrame([dict(r) for r in monthly])
        df_monthly = df_monthly.rename(columns={'month': '月份', 'count': '筆數', 'total_amount': '金額'})
        st.line_chart(df_monthly.set_index('月份')['金額'])
    else:
        st.info("尚無出貨記錄。")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 📋 訂單管理
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
elif page == "📋 訂單管理":
    st.title("📋 訂單管理")

    # 篩選列
    filter_col1, filter_col2 = st.columns(2)
    with filter_col1:
        filter_status = st.selectbox(
            "依狀態篩選",
            ["全部", "製程中", "已出貨", "客戶暫停", "待零件", "延遲", "已取消"],
        )
    with filter_col2:
        filter_part = st.text_input("搜尋料號", placeholder="輸入料號關鍵字...")

    status_param = None if filter_status == "全部" else filter_status
    part_param = filter_part if filter_part else None
    orders = db.get_orders(status=status_param, part_no=part_param)

    if orders:
        df = pd.DataFrame([dict(r) for r in orders])
        display_cols = ['id', 'order_no', 'part_no', 'quantity', 'amount', 'vendor_name',
                       'due_date', 'status', 'source_sheet', 'note']
        display_cols = [c for c in display_cols if c in df.columns]
        df_display = df[display_cols].rename(columns={
            'id': 'ID', 'order_no': '訂單號', 'part_no': '料號',
            'quantity': '數量', 'amount': '金額', 'vendor_name': '廠商',
            'due_date': '預定交期', 'status': '狀態',
            'source_sheet': '來源', 'note': '備註'
        })

        # 顏色標記
        def color_status(val):
            colors = {
                '製程中': 'background-color: #fff3cd',
                '已出貨': 'background-color: #d4edda',
                '延遲': 'background-color: #f8d7da',
                '客戶暫停': 'background-color: #d6d8db',
                '待零件': 'background-color: #cce5ff',
            }
            return colors.get(val, '')

        styled = df_display.style.map(color_status, subset=['狀態'])
        st.dataframe(styled, hide_index=True, use_container_width=True, height=500)
        st.caption(f"共 {len(df_display)} 筆訂單")
    else:
        st.info("沒有符合條件的訂單。")

    # 新增訂單
    st.divider()
    with st.expander("➕ 新增訂單"):
        with st.form("new_order"):
            nc1, nc2 = st.columns(2)
            with nc1:
                new_part = st.text_input("料號 *", placeholder="例：YB267A010D")
                new_qty = st.number_input("數量", min_value=0, value=0)
                new_vendor = st.text_input("廠商")
            with nc2:
                new_order_no = st.text_input("訂單號")
                new_amount = st.number_input("金額", min_value=0.0, value=0.0)
                new_due = st.date_input("預定交期")

            new_note = st.text_input("備註")
            submitted = st.form_submit_button("建立訂單", type="primary")

            if submitted and new_part:
                db.insert_order(
                    order_no=new_order_no or None,
                    part_no=new_part,
                    quantity=new_qty if new_qty > 0 else None,
                    amount=new_amount if new_amount > 0 else None,
                    vendor_name=new_vendor or None,
                    due_date=str(new_due),
                    note=new_note or None,
                )
                st.success(f"✓ 訂單已建立: {new_part}")
                st.rerun()
            elif submitted:
                st.warning("請填寫料號。")

    # 更新狀態
    with st.expander("✏️ 更新訂單狀態"):
        with st.form("update_order"):
            uo_id = st.number_input("訂單 ID", min_value=1, step=1)
            uo_status = st.selectbox("新狀態", ["製程中", "已出貨", "客戶暫停", "待零件", "延遲", "已取消"])
            uo_note = st.text_input("備註更新")
            uo_submit = st.form_submit_button("更新", type="primary")
            if uo_submit:
                kwargs = {"status": uo_status}
                if uo_note:
                    kwargs["note"] = uo_note
                db.update_order(uo_id, **kwargs)
                st.success(f"✓ 訂單 #{uo_id} 已更新為「{uo_status}」")
                st.rerun()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 🔧 製程追蹤
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
elif page == "🔧 製程追蹤":
    st.title("🔧 製程追蹤")

    search_part = st.text_input("🔍 搜尋料號", placeholder="輸入料號查看製程進度...")

    if search_part:
        steps = db.get_process_steps(part_no=search_part)
        if steps:
            df_steps = pd.DataFrame([dict(r) for r in steps])

            # 顯示進度條
            unique_parts = df_steps['part_no'].unique()
            for pn in unique_parts:
                pn_steps = df_steps[df_steps['part_no'] == pn]
                total = len(pn_steps)
                completed = len(pn_steps[pn_steps['status'] == '完成'])
                pct = completed / total if total > 0 else 0

                st.subheader(f"📦 {pn}")
                st.progress(pct, text=f"完成 {completed}/{total} 道製程 ({pct:.0%})")

                # 製程明細表
                display_steps = pn_steps[['step_seq', 'process_name', 'vendor_name',
                                          'planned_date', 'actual_date', 'status', 'work_date']].rename(columns={
                    'step_seq': '序號', 'process_name': '製程',
                    'vendor_name': '廠商', 'planned_date': '預定日期',
                    'actual_date': '實際日期', 'status': '狀態', 'work_date': '排程日'
                })

                def color_step_status(val):
                    colors = {
                        '完成': 'background-color: #d4edda; color: #155724',
                        '進行中': 'background-color: #fff3cd; color: #856404',
                        '延遲': 'background-color: #f8d7da; color: #721c24',
                        '待處理': 'background-color: #e2e3e5; color: #383d41',
                    }
                    return colors.get(val, '')

                styled_steps = display_steps.style.map(color_step_status, subset=['狀態'])
                st.dataframe(styled_steps, hide_index=True, use_container_width=True)
                st.divider()
        else:
            st.warning(f"找不到料號「{search_part}」的製程資料。")
    else:
        # 顯示所有料號的進度總覽
        st.subheader("全部料號進度總覽")
        all_steps = db.get_process_steps()
        if all_steps:
            df_all = pd.DataFrame([dict(r) for r in all_steps])
            summary = df_all.groupby('part_no').agg(
                total=('status', 'count'),
                completed=('status', lambda x: (x == '完成').sum()),
                in_progress=('status', lambda x: (x == '進行中').sum()),
            ).reset_index()
            summary['完成率'] = (summary['completed'] / summary['total'] * 100).round(1)
            summary = summary.rename(columns={
                'part_no': '料號', 'total': '總步驟',
                'completed': '已完成', 'in_progress': '進行中'
            })
            summary = summary.sort_values('完成率', ascending=False)
            st.dataframe(summary, hide_index=True, use_container_width=True, height=500)
        else:
            st.info("尚無製程資料。")

    # 更新製程狀態
    st.divider()
    with st.expander("✏️ 更新製程步驟"):
        with st.form("update_step"):
            us_col1, us_col2 = st.columns(2)
            with us_col1:
                us_id = st.number_input("製程步驟 ID", min_value=1, step=1)
                us_status = st.selectbox("新狀態", ["待處理", "進行中", "完成", "延遲", "跳過"])
            with us_col2:
                us_actual = st.date_input("實際完成日期")
                us_note = st.text_input("備註")
            us_submit = st.form_submit_button("更新製程", type="primary")
            if us_submit:
                kwargs = {"status": us_status, "actual_date": str(us_actual)}
                if us_note:
                    kwargs["note"] = us_note
                db.update_process_step(us_id, **kwargs)
                st.success(f"✓ 製程步驟 #{us_id} 已更新")
                st.rerun()

    # 新增製程步驟
    with st.expander("➕ 新增製程步驟"):
        with st.form("new_step"):
            ns_col1, ns_col2 = st.columns(2)
            with ns_col1:
                ns_order_id = st.number_input("訂單 ID", min_value=1, step=1)
                ns_part = st.text_input("料號 *")
                ns_seq = st.number_input("步驟序號", min_value=1, step=1)
            with ns_col2:
                ns_process = st.text_input("製程名稱 *", placeholder="例：裁切、CNC、PTH...")
                ns_vendor = st.text_input("廠商")
                ns_planned = st.date_input("預定日期")
            ns_submit = st.form_submit_button("新增", type="primary")
            if ns_submit and ns_part and ns_process:
                db.insert_process_step(
                    order_id=ns_order_id,
                    part_no=ns_part,
                    step_seq=ns_seq,
                    process_name=ns_process,
                    vendor_name=ns_vendor or None,
                    planned_date=str(ns_planned),
                    status="待處理",
                )
                st.success(f"✓ 已新增製程步驟: {ns_process}")
                st.rerun()
            elif ns_submit:
                st.warning("請填寫料號和製程名稱。")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 🚚 出貨排程
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
elif page == "🚚 出貨排程":
    st.title("🚚 出貨排程")

    tab1, tab2 = st.tabs(["📦 出貨記錄", "➕ 新增出貨"])

    with tab1:
        ship_col1, ship_col2 = st.columns(2)
        with ship_col1:
            ship_part = st.text_input("搜尋料號", key="ship_search")
        with ship_col2:
            ship_month = st.text_input("篩選月份 (YYYY-MM)", placeholder="例：2026-03")

        shipments = db.get_shipments(
            part_no=ship_part if ship_part else None,
            month=ship_month if ship_month else None,
        )

        if shipments:
            df_ship = pd.DataFrame([dict(r) for r in shipments])
            display_ship = df_ship[['id', 'part_no', 'ship_date', 'ship_quantity', 'amount', 'note']].rename(columns={
                'id': 'ID', 'part_no': '料號', 'ship_date': '出貨日期',
                'ship_quantity': '出貨數量', 'amount': '金額', 'note': '備註'
            })
            st.dataframe(display_ship, hide_index=True, use_container_width=True, height=400)

            total_amount = df_ship['amount'].sum()
            total_qty = df_ship['ship_quantity'].sum()
            m1, m2, m3 = st.columns(3)
            m1.metric("筆數", len(df_ship))
            m2.metric("總數量", f"{total_qty:,.0f}" if pd.notna(total_qty) else "N/A")
            m3.metric("總金額", f"${total_amount:,.0f}" if pd.notna(total_amount) else "N/A")
        else:
            st.info("沒有符合條件的出貨記錄。")

    with tab2:
        with st.form("new_shipment"):
            sc1, sc2 = st.columns(2)
            with sc1:
                ns_part = st.text_input("料號 *", key="new_ship_part")
                ns_qty = st.number_input("出貨數量", min_value=0, step=1)
                ns_date = st.date_input("出貨日期")
            with sc2:
                ns_order_id = st.number_input("關聯訂單 ID (選填)", min_value=0, step=1)
                ns_amount = st.number_input("金額", min_value=0.0, value=0.0)
                ns_note = st.text_input("備註")

            ns_submit = st.form_submit_button("登記出貨", type="primary")
            if ns_submit and ns_part:
                db.insert_shipment(
                    order_id=ns_order_id if ns_order_id > 0 else None,
                    part_no=ns_part,
                    ship_date=str(ns_date),
                    ship_quantity=ns_qty if ns_qty > 0 else None,
                    amount=ns_amount if ns_amount > 0 else None,
                    note=ns_note or None,
                )
                # 同時更新訂單狀態
                if ns_order_id > 0:
                    db.update_order(ns_order_id, status="已出貨")
                st.success(f"✓ 出貨已登記: {ns_part}")
                st.rerun()
            elif ns_submit:
                st.warning("請填寫料號。")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 🏗️ 廠商統計
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
elif page == "🏗️ 廠商統計":
    st.title("🏗️ 廠商統計")

    tab1, tab2 = st.tabs(["📊 負載統計", "📝 廠商清單"])

    with tab1:
        vendor_load = db.get_vendor_load()
        if vendor_load:
            df_vl = pd.DataFrame([dict(r) for r in vendor_load])
            df_vl = df_vl.rename(columns={
                'vendor_name': '廠商', 'total_steps': '總製程數',
                'completed': '已完成', 'in_progress': '進行中', 'delayed': '延遲'
            })
            df_vl['完成率'] = (df_vl['已完成'] / df_vl['總製程數'] * 100).round(1)

            st.dataframe(df_vl, hide_index=True, use_container_width=True)

            # 廠商負載長條圖
            st.subheader("廠商製程負載")
            chart_data = df_vl.set_index('廠商')[['已完成', '進行中', '延遲']]
            st.bar_chart(chart_data)
        else:
            st.info("尚無廠商製程資料。")

    with tab2:
        vendors = db.get_vendors()
        if vendors:
            df_vendors = pd.DataFrame([dict(r) for r in vendors])
            df_vendors = df_vendors[['id', 'name', 'contact', 'phone', 'note']].rename(columns={
                'id': 'ID', 'name': '廠商名稱', 'contact': '聯絡人',
                'phone': '電話', 'note': '備註'
            })
            st.dataframe(df_vendors, hide_index=True, use_container_width=True)
        else:
            st.info("尚無廠商資料。")

        # 新增廠商
        with st.expander("➕ 新增/更新廠商"):
            with st.form("new_vendor"):
                nv1, nv2 = st.columns(2)
                with nv1:
                    nv_name = st.text_input("廠商名稱 *")
                    nv_contact = st.text_input("聯絡人")
                with nv2:
                    nv_phone = st.text_input("電話")
                    nv_note = st.text_input("備註")
                nv_submit = st.form_submit_button("儲存", type="primary")
                if nv_submit and nv_name:
                    db.upsert_vendor(nv_name, nv_contact, nv_phone, nv_note)
                    st.success(f"✓ 廠商已儲存: {nv_name}")
                    st.rerun()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 📥 資料匯入
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
elif page == "📥 資料匯入":
    st.title("📥 資料匯入")
    st.markdown("""
    從現有的 Excel (.xls) 檔案匯入資料到系統中。

    **支援檔案：**
    - `2026.03急件.xls` → 匯入製程追蹤資料
    - `2026當周出貨排程.xls` → 匯入訂單與出貨記錄
    """)

    import import_xls

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📄 製程追蹤表")
        file1_exists = os.path.exists(import_xls.FILE1)
        st.text(f"檔案: {'✓ 已找到' if file1_exists else '✗ 未找到'}")
        if file1_exists:
            if st.button("匯入製程追蹤", type="primary"):
                with st.spinner("匯入中..."):
                    count = import_xls.import_file1_process_tracking()
                st.success(f"✓ 已匯入 {count} 筆製程步驟")

    with col2:
        st.subheader("📄 出貨排程")
        file2_exists = os.path.exists(import_xls.FILE2)
        st.text(f"檔案: {'✓ 已找到' if file2_exists else '✗ 未找到'}")
        if file2_exists:
            if st.button("匯入出貨排程", type="primary"):
                with st.spinner("匯入中..."):
                    result = import_xls.import_file2_shipping()
                if result:
                    st.success(f"✓ 已匯入 {result[0]} 筆訂單, {result[1]} 筆出貨記錄")

    st.divider()
    if st.button("🔄 完整重新匯入（清空後匯入）", type="secondary"):
        with st.spinner("清空資料庫並重新匯入..."):
            # 清空所有表
            with db.get_conn() as conn:
                conn.execute("DELETE FROM process_steps")
                conn.execute("DELETE FROM shipments")
                conn.execute("DELETE FROM orders")
                conn.execute("DELETE FROM vendors")
            import_xls.run_full_import()
        st.success("✓ 已完成完整重新匯入")
        st.rerun()

    # 顯示目前資料庫統計
    st.divider()
    st.subheader("📊 目前資料庫統計")
    stats = db.get_dashboard_stats()
    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.metric("訂單", stats['total_orders'])
    sc2.metric("製程中", stats['in_progress'])
    sc3.metric("已出貨", stats['shipped'])
    sc4.metric("延遲", stats['delayed'])


# ── Footer ──
st.sidebar.divider()
st.sidebar.caption("妍發科技 製程管理系統 v1.0")
st.sidebar.caption("Powered by Streamlit + SQLite")
