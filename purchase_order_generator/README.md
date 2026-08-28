# 📦 采购单自动生成系统 (Purchase Order Generator)

朗晨 POPYOUNG 品牌采购单批量生成工具：输入一个 Excel（3 个工作表），自动按 **款号×工厂** 拆分生成多份采购单（生产订单 + 每布行布料申购单），并额外生成**领星 ERP 批量导入采购单**，全部打包 zip 下载。

## ✨ 功能

- ✅ 按 `款号 × 工厂` 自动拆单，订单号 `PPY+日期+序号` 自动生成
- ✅ 生产订单（THD 模板样式）：颜色×尺码动态矩阵、GCD 比例、布料条数、汇总行、图片内置
- ✅ 布料申购单：**每个布行一个独立 sheet**，条数自动换算（件数÷每条布出货数，向上取整）
- ✅ **混搭色**（如 `#179 海军蓝+YH5126 棕蓝格子`）自动拆两行分归不同布行，预览可改拆分数量
- ✅ 7 条数据校验（同款同色工厂/布行一致性防复制错误等）
- ✅ 领星批量导入采购单生成（标识号分组、供应商自动加"工厂"后缀）
- ✅ 汇总明细表 + zip 打包下载
- ✅ 输入模板 + 样本文件一键下载

## 🚀 快速开始

```bash
git clone <your-repo-url>
cd purchase-order-generator
python -m venv .venv

# Windows
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python -m streamlit run app.py --server.headless true --server.port 8501
```

打开 http://127.0.0.1:8501

## 📥 输入文件格式（3 个工作表）

| Sheet | 内容 |
|---|---|
| **主表** | 款号 \| SKU编码 \| SKU名称 \| 颜色 \| 尺码 \| 采购数 \| 工厂 \| 布行 \| 品名 \| 备注 |
| **面料表** | 款号 \| 产品名称 \| 品名 \| 纸样名称 \| 洗水唛成分 \| 每条布出货数(件) \| 是否需要压缩袋 \| 包装袋规格 |
| **配置表** | 品牌 \| 账号代码 \| 申请人 \| 联系地址 \| 采购仓库 |

> 从应用首页下载**样本输入文件**或运行 `python make_input_template.py` 生成模板。
> 输入文件名需含 8 位日期：`PPY采购-YYYYMMDD.xlsx`（用于生成订单号）。

## 📂 目录结构

```
├── app.py                  # Streamlit 应用
├── core/
│   ├── calc.py             # 计算工具（比例/条数/尺码/混搭/布行分隔）
│   ├── parser.py           # 输入解析 + 校验
│   ├── splitter.py         # 拆单逻辑
│   ├── generator.py        # 输出 xlsx 生成
│   └── lingxing.py         # 领星导入文件生成
├── templates/              # 模板文件（见下方隐私说明）
├── make_input_template.py  # 输入模板生成
├── make_sample_input.py    # 样本输入生成
├── make_clean_template.py  # 净化模板生成
└── e2e_test.py             # 端到端测试
```

## ⚠️ 隐私说明

`templates/` 下的模板文件（THD 模板、领星模板）**包含公司业务数据**（品牌、仓库名、工厂名、订单格式）。
- 建议使用**私有仓库**部署
- 或将 `templates/*.xlsx` 加入 `.gitignore`，部署后自行放置模板

模板依赖：
- `templates/thd_template_clean.xlsx` — 生产订单底版（图片已内置）
- `templates/ppy_old_template.xlsx` — 布料申购单样式来源
- `templates/lingxing_template.xlsx` — 领星批量导入模板底版
- `templates/images/` — 4 张内置图片

## 🧪 测试

```bash
python e2e_test.py   # 解析→拆单→生成 全链路
```

## 📄 更多文档

- `使用手册.md` — 面向使用者的详细操作说明（含输入文件填写规范）
- `HANDOFF.md` — 面向开发/AI 接手的架构与业务规则说明
