from pathlib import Path

src = Path(__file__).with_name("app_base.py").read_text(encoding="utf-8")

def repl(old: str, new: str, count=None):
    global src
    if old not in src:
        raise RuntimeError(f"V3.1 patch target not found: {old[:80]}")
    src = src.replace(old, new) if count is None else src.replace(old, new, count)

repl('RULE_VERSION = "蒙西2026（自动结算V2.9）"', 'RULE_VERSION = "蒙西2026（自动结算V3.1）"')
repl('st.sidebar.caption("Supabase 云数据库版 · 自动结算V2.9")', 'st.sidebar.caption("Supabase 云数据库版 · 自动结算V3.1")')

repl('''    mp = float(inp.get("mechanism_price") or BASELINE_PRICE)\n    ms = float(inp.get("mechanism_spot_price") or 0.0)\n    mechanism = me * (mp - ms) if me > 0 and ms != 0 else 0.0\n    regular = float(inp.get("regular_fee") or 0.0)\n    unit = float(inp.get("unit_fee") or 0.0)\n    manual = float(inp.get("manual_adjustment") or 0.0)\n    final = energy + congestion - assessment + risk + green_fee + mechanism + regular - unit + manual''', '''    mechanism = float(inp.get("mechanism_fee_manual") or 0.0)\n    replacement = float(inp.get("replacement_fee") or 0.0)\n    regular = float(inp.get("regular_fee") or 0.0)\n    unit = float(inp.get("unit_fee") or 0.0)\n    manual = float(inp.get("manual_adjustment") or 0.0)\n    settlement_revenue = energy - assessment + risk + green_fee + mechanism + replacement + regular - unit + manual\n    final = settlement_revenue + congestion''')

repl('''        "green_fee": green_fee,\n        "mechanism_fee": mechanism,\n        "final_revenue": final,\n        "final_price": final / act if act else 0.0,''', '''        "green_fee": green_fee,\n        "mechanism_fee": mechanism,\n        "replacement_fee": replacement,\n        "settlement_revenue": settlement_revenue,\n        "final_revenue": final,\n        "final_price": settlement_revenue / act if act else 0.0,''')

repl('''                a, b, c = st.columns(3)\n                me = a.number_input("机制电量 MWh", value=float(old.get("mechanism_energy") or 0))\n                mp = b.number_input("机制电价", value=float(old.get("mechanism_price") or BASELINE_PRICE))\n                ms = c.number_input("机制电量对应现货均价", value=float(old.get("mechanism_spot_price") or 0))''', '''                a, b, c = st.columns(3)\n                me = a.number_input("机制电量 MWh", value=float(old.get("mechanism_energy") or 0))\n                mechanism_fee_manual = b.number_input("机制费用（手工填写） 元", value=float(old.get("mechanism_fee_manual") or 0))\n                replacement_fee = c.number_input("置换费用（手工填写） 元", value=float(old.get("replacement_fee") or 0))''')

repl('''                        "green_environment_price": gp, "mechanism_energy": me, "mechanism_price": mp,\n                        "mechanism_spot_price": ms, "unit_fee": unit, "manual_adjustment": manual,''', '''                        "green_environment_price": gp, "mechanism_energy": me,\n                        "mechanism_fee_manual": mechanism_fee_manual, "replacement_fee": replacement_fee,\n                        "unit_fee": unit, "manual_adjustment": manual,''')

repl('''            ["机制", "机制费用", "Q机制×(P机制-P机制现货)", calc["mechanism_fee"], "元"],\n            ["结果", "预计最终收益", "自动汇总", calc["final_revenue"], "元"],\n            ["结果", "最终结算均价", "最终收益/Qactual", calc["final_price"], "元/MWh"],''', '''            ["机制", "机制费用", "手工填写", calc["mechanism_fee"], "元"],\n            ["置换", "置换费用", "手工填写", calc["replacement_fee"], "元"],\n            ["结果", "结算口径收入", "不含阻塞盈余，含置换费用", calc["settlement_revenue"], "元"],\n            ["结果", "含阻塞总收益", "结算口径收入+阻塞盈余", calc["final_revenue"], "元"],\n            ["结果", "最终结算均价", "结算口径收入/Qactual（不含阻塞、含置换）", calc["final_price"], "元/MWh"],''')

repl('''    inp = load_input(int(r.id), month)\n    calc = auto_settlement(sm, inp)''', '''    inp = load_input(int(r.id), month)\n    form_scope = f"{int(r.id)}_{month}"\n    with st.expander("费用手工录入｜机制费用 / 置换费用", expanded=False):\n        with st.form(f"station_statement_manual_fees_{form_scope}"):\n            fa, fb, fc = st.columns(3)\n            fee_me = fa.number_input("机制电量 MWh", value=float(inp.get("mechanism_energy") or 0), key=f"report_me_{form_scope}")\n            fee_mech = fb.number_input("机制费用（手工填写） 元", value=float(inp.get("mechanism_fee_manual") or 0), key=f"report_mech_fee_{form_scope}")\n            fee_replace = fc.number_input("置换费用（手工填写） 元", value=float(inp.get("replacement_fee") or 0), key=f"report_replace_fee_{form_scope}")\n            if st.form_submit_button("保存费用并重算", type="primary"):\n                save_input(int(r.id), month, {\n                    "mechanism_energy": fee_me,\n                    "mechanism_fee_manual": fee_mech,\n                    "replacement_fee": fee_replace,\n                })\n                st.rerun()\n    inp = load_input(int(r.id), month)\n    calc = auto_settlement(sm, inp)''', count=1)

repl('''            ["阻塞盈余", float(inp.get("congestion") or 0.0), "元"],\n            ["风险防范", calc["risk_prevention"], "元"],\n            ["绿电费用", calc["green_fee"], "元"],\n            ["机制费用", calc["mechanism_fee"], "元"],\n            ["机组/两个细则扣费", -float(inp.get("unit_fee") or 0.0), "元"],\n            ["预计最终收益", calc["final_revenue"], "元"],\n            ["最终结算均价", calc["final_price"], "元/MWh"],''', '''            ["阻塞盈余（不计入最终均价）", float(inp.get("congestion") or 0.0), "元"],\n            ["风险防范", calc["risk_prevention"], "元"],\n            ["绿电费用", calc["green_fee"], "元"],\n            ["机制费用（手工）", calc["mechanism_fee"], "元"],\n            ["置换费用（手工）", calc["replacement_fee"], "元"],\n            ["机组/两个细则扣费", -float(inp.get("unit_fee") or 0.0), "元"],\n            ["结算口径收入（不含阻塞）", calc["settlement_revenue"], "元"],\n            ["含阻塞总收益", calc["final_revenue"], "元"],\n            ["最终结算均价", calc["final_price"], "元/MWh"],''')

# 自动结算页切换场站/月份时，也重新创建表单，避免沿用上一场站输入状态。
repl('''            with st.form("params"):''', '''            with st.form(f"params_{int(r.id)}_{month}"):''')

# Excel 风格数据条：价格/电费直接在单元格背景里显示，不额外增加趋势列。
repl('''    st.dataframe(statement_show, use_container_width=True, hide_index=True, height=760)''', '''    statement_style = (\n        statement_show.style\n        .bar(subset=["日现货均价", "现货全电费"], color="#63BE7B", align="zero")\n        .bar(subset=["中长期差价电费"], color=["#F8696B", "#63BE7B"], align="mid")\n        .bar(subset=["电能电费"], color="#5B9BD5", align="zero")\n        .format(precision=4, subset=numeric_cols)\n    )\n    st.dataframe(statement_style, use_container_width=True, hide_index=True, height=760)''')

# 费用与均价使用静态表格，一次性展示全部行，不出现表格内部纵向滚动条。
repl('''        st.dataframe(fee_summary, use_container_width=True, hide_index=True)''', '''        st.table(fee_summary)''')

repl('''    st.write("场站月报：按整月日期展示合约电量/均价、上网电量、日现货均价、现货电费、中长期差价电费、电能电费，并自动汇总结算项")''', '''    st.write("场站月报：按场站+月份独立保存机制费用/置换费用；费用与均价完整展开；月报单元格显示Excel风格数据条；最终结算均价不计阻塞盈余、计入置换费用")''')

exec(compile(src, str(Path(__file__).with_name("app_base.py")), "exec"), globals(), globals())
