# 采购单自动生成系统 — AI 交接文档（HANDOFF）

> 本文档供**其他 AI 代理接手本项目**时快速理解全部上下文。
> 项目版本：v1.5 | 最后更新：2026-08-27 | 维护人：朗晨（POPYOUNG）

---

## 1. 项目是什么

**朗晨 PPY 品牌采购单批量生成工具**。运营算好每个 SKU 采购数量后，上传一个 Excel 输入文件，系统自动：
1. 按 **款号 × 工厂** 拆分 → 生成 N 份采购单（生产订单 + 每布行一张布料申购单）
2. 额外生成**领星 ERP 批量导入采购单**（对接领星系统）
3. 生成**汇总明细表**，全部打包 zip 下载

**核心痛点解决**：一次采购多款、每款拆多工厂、每款涉及多布行 → 手工做多份文件易错 → 自动化。

## 2. 工程结构

```
purchase-order-generator/
├── app.py                  # Streamlit 界面（上传→校验→预览→生成→zip下载）
├── requirements.txt        # streamlit / pandas / openpyxl
├── make_sample_input.py    # 生成样例输入文件（测试用）
├── make_input_template.py  # 生成输入模板（3 工作表）
├── make_clean_template.py  # 生成净化模板（把 DISPIMG 图换成普通浮动图）
├── e2e_test.py             # 端到端测试脚本
├── verify_output.py        # 输出文件验证脚本
├── core/
│   ├── calc.py             # 工具：GCD比例/条数/尺码排序/类型/混搭色/布行分隔符
│   ├── parser.py           # 输入解析 + 7 条校验 + 图片解析
│   ├── splitter.py         # 拆单（款×工厂/订单号/布行聚合/混搭拆分）
│   ├── generator.py        # 输出 xlsx 生成（订单 sheet + 申购单 sheet + 汇总表）
│   └── lingxing.py         # 领星批量导入采购单生成
├── templates/
│   ├── thd_template.xlsx         # 原始 THD 模板（含 DISPIMG 图片，勿直接用作底版）
│   ├── thd_template_clean.xlsx   # ★净化模板（generator 底版，图片已内置为普通浮动）
│   ├── ppy_old_template.xlsx     # PPY 旧模板（布料申购单样式来源）
│   ├── lingxing_template.xlsx    # ★领星批量导入模板底版
│   └── images/                   # 4 张内置默认图（包装袋/主唛/吊牌/卡片）
└── output/                 # 生成物（zip、样例、模板）
```

## 3. 数据流

```
输入文件(1个xlsx, 3个sheet)
  ├─ Sheet1 主表(10列): 款号|SKU编码|SKU名称|颜色|尺码|采购数|工厂|布行|品名|备注
  ├─ Sheet2 面料表(8列): 款号|产品名称|品名|纸样名称|洗水唛成分|每条布出货数(件)|是否需要压缩袋|包装袋规格
  └─ Sheet3 配置表: 品牌|账号代码|申请人|联系地址|采购仓库
        ↓ parser.py (校验)
        ↓ splitter.py (拆单)
  输出 (zip打包):
  ├─ {订单号}-{款号} {产品名称} {类型} -{工厂}.xlsx   ← 每个 款×工厂 一份
  │    ├─ Sheet「生产订单」(THD 模板 A-L 列样式, 含图片)
  │    └─ Sheet「布料申购单-{布行}」× N  (每布行一个)
  ├─ {日期}_采购单汇总.xlsx      (文件汇总 + 布行明细)
  └─ {日期}_领星批量导入采购单.xlsx
```

## 4. 核心业务规则（改代码前必读）

| 规则 | 说明 | 代码位置 |
|---|---|---|
| **订单号** | `PPY` + 输入文件名日期 + 序号；序号按 **款×工厂组合** 首次出现顺序 01/02/03 | splitter.py |
| **文件命名** | `{订单号}-{款号} {产品名称} {类型} -{工厂}.xlsx`；类型自动拼（纯色/印花，**混搭算印花**） | splitter.py |
| **布料条数** | 件数 ÷ 每条布出货数，**向上取整**；订单 sheet 条数 = 各布行聚合后条数合计（文件内自洽） | calc.rolls / splitter |
| **混搭色** | 颜色含 `+`（如 `#179 海军蓝+YH5126 棕蓝格子`）：订单 sheet 一行原样；申购单**拆两行**（布行按逗号/空格/分号拆 2 个，前半归第1个布行、后半归第2个；采购数对半，可 override） | calc.split_suppliers / splitter |
| **尺码** | 动态 S~6XL，只列有采购数的尺码，按标准序排列 | calc.SIZE_ORDER |
| **比例列** | 各尺码数 ÷ 最大公约数（GCD）自动算 | calc.gcd_ratio |
| **校验** | 同款同色 SKU 的工厂/布行必须一致（防复制错误）；混搭布行必须 2 个；SKU 唯一；尺码合法 | parser.validate |
| **订单 sheet 图片** | 用 thd_template_clean 底版（图已内置）；插行时 `_shift_images_down` 让图片随固定内容下移 | generator |
| **领星文件** | 标识号=款×工厂顺序；供应商+“工厂”后缀；填 含税否/不分配/CNY/采购仓库/SKU/实际采购量；**其余留空**（模板示例留空的不填） | lingxing.py |

## 5. 运行方式

```bash
# 环境（Windows）
cd purchase-order-generator
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python -m streamlit run app.py --server.headless true --server.port 8501
# 浏览器打开 http://127.0.0.1:8501
```

## 6. 模板说明（重要）

- **thd_template_clean.xlsx**：generator 底版。THD 原模板的 DISPIMG（单元格嵌入图片）openpyxl 无法保留，已用 make_clean_template.py 替换成普通浮动图片。用户以后想改图片：直接打开此模板改图（普通更换，**不要用“置于单元格”**）。
- **lingxing_template.xlsx**：领星批量导入模板底版，复制自 `F:\领星项目\采购\批量导入采购单模板-V376`。
- **templates/images/**：4 张内置默认图，DISPIMG_DEFAULTS 里配置了紧凑尺寸（包装袋 350×230 / 主唛 320×240 / 吊牌 260×110 / 卡片 300×250）——**模板原 DISPIMG ext 过大（主唛 783×584）会导致图片超出视觉区，别改回大尺寸**。

## 7. 已知问题 / 注意事项

1. ⚠️ **openpyxl 不支持 DISPIMG**：任何含“置于单元格”图片的 xlsx 经 openpyxl 保存都会丢图。遇到图片问题优先走“普通浮动图片”方案。
2. ⚠️ **Streamlit 缓存**：修改 core/*.py 后必须**重启 streamlit 进程**（sys.modules 缓存旧模块）。
3. ⚠️ **模板含业务数据**（品牌/仓库名/工厂名）：部署 GitHub 建议**私有仓库**，templates/ 可考虑 .gitignore。
4. 空行/汇总行的 B、L 列必须**显式写空**（THD 模板有 SUMIFS/压缩袋公式残留，否则 #REF!）。
5. 重读 openpyxl 图片 `img.width` 返回的是 **PIL 原始尺寸**，实际渲染看 `anchor.ext.cx/cy`（EMU，÷9525 转 px）。
6. 混搭色拆分数量程序默认对半（奇数前段多 1），预览界面可改。

## 8. 后续可做

- 把图片方案重新启用（输入文件第 4 sheet「图片」的解析代码在 parser 里保留着，模板已移除）
- 部署到阿里云 OSS / 服务器（当前本地运行）
- 接领星 API 直接建单（当前是文件导入）
- 老格式输入（每款一 sheet）自动转新表头脚本
