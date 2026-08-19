# 蒙西新能源多场站日清分与交易决策系统

面向蒙西发电侧新能源交易的 Streamlit 工具。支持多场站管理、批量上传日清分 Excel、96 点自动识别与 24 时点计算、日汇总、月度累计、结算费用以及 P10/P50/P90 仓位交易决策。

## 当前架构

```text
Streamlit
   ↓
Supabase Data API
   ↓
PostgreSQL
   ├─ stations        场站档案
   ├─ daily_summary   每日清分“总”与结算费用
   └─ hourly_detail   每天24时点明细
```

系统已经从本地 SQLite 迁移为 Supabase 云数据库。同一数据库可被不同电脑/部署实例共享。

## 功能

- 自助创建光伏/风电场站，并配置简称、装机容量和仓位风控参数
- 场站信息可编辑、启用/停用
- 一次上传多个日清分 `.xlsx/.xlsm`
- 根据场站名称/简称自动匹配场站，无法识别时可手工选择
- 自动识别交易日期和 96 点原始数据
- 96 个 15 分钟点转换为 24 个时点并计算日“总”
- 同场站 + 同日期唯一约束，防止重复累计；可选择覆盖重算
- 按场站、月份累计电量、加权均价、收益、考核、风险防范等
- 单日录入阻塞盈余、考核、绿电、机制、风险防范、机组费用等结算项
- 决策模型直接读取真实累计结果，再叠加剩余 P10/P50/P90 发电预测
- 输出低发场景仓位、保守可售空间、建议卖出/买回量及交易动作
- 数据管理支持查看历史日清分、24时点及删除错误数据

## 当前日清分模板

当前解析器按已确认的蒙西日清分模板读取：

- C：中长期电力
- D：中长期价格
- J：实时价格
- K：实际计量电力
- L：全网统一结算点价格

96 点按每 4 个 15 分钟点合成 1 个小时点。

小时计算：

- 中长期电量 = C 的 4 点平均
- 中长期价 = D 的 4 点平均
- 实际上网电量 = K 的 4 点平均
- 实时价 = J 按 K 加权；K 合计为 0 时使用 J 平均
- 统一结算点价 = L 的 4 点平均
- 现货电费 = 实际上网电量 × 实时价
- 中长期差价电费 = 中长期电量 × (中长期价 - 统一结算点价)
- 电能量合计 = 现货电费 + 中长期差价电费

月度价格采用电量加权，不对每日均价做简单算术平均。

## 首次连接 Supabase

### 1. 初始化数据库

在 Supabase 项目中进入 **SQL Editor**，打开仓库中的 `supabase_schema.sql`，完整执行一次。

该 SQL 会创建：

- `stations`
- `daily_summary`
- `hourly_detail`
- 唯一约束和索引
- V1 所需 RLS policy

### 2. 配置密钥

真实密钥 **不要写入 GitHub**。

本地创建：

```text
.streamlit/secrets.toml
```

内容：

```toml
SUPABASE_URL = "https://bwhpatzqcklgyzuifvhu.supabase.co"
SUPABASE_KEY = "你的 publishable key"
```

`.gitignore` 已忽略 `.streamlit/secrets.toml`。

如果部署到 Streamlit Community Cloud，在应用的 **Settings → Secrets** 中填写相同两项即可，不需要把 secrets 文件提交到仓库。

## 本地运行

```bash
git clone https://github.com/veen-one/mengxi-power-trading-system.git
cd mengxi-power-trading-system
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Windows 激活虚拟环境：

```powershell
.venv\Scripts\activate
```

## 文件说明

```text
app.py                         Streamlit 主程序
supabase_backend.py            Supabase 数据读写封装
supabase_schema.sql            数据库初始化 SQL
requirements.txt               Python 依赖
.streamlit/secrets.toml.example Secrets 示例（无真实 key）
.gitignore                     防止本地密钥/缓存提交
```

## 结算口径

当前系统已实现日清分、电能量合计和扩展结算费用的累计框架。考核、风险防范、曲线合理度、机制电价等正式算法需要按对应交易月份的正式规则参数维护。系统中的交易决策属于交易风控/压力测试，不替代交易中心最终结算结果。

## 安全说明

- 仓库中不保存真实 Supabase key。
- `SUPABASE_KEY` 仅从 Streamlit Secrets 或系统环境变量读取。
- 页面不会显示 key。
- 当前 `supabase_schema.sql` 的 V1 policy 用于个人快速使用；若正式公网部署并开放给多人使用，建议下一阶段接入 Supabase Auth，并按用户/角色收紧 RLS。
