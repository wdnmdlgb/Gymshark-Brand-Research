# Gymshark 品牌深度研究 📊

> 桌面研究项目 | 欧洲健身服饰DTC品牌竞品分析与出海策略参考
>
> 个人求职作品集项目，非官方研究，仅用于学习目的

---

## 📖 快速预览

| 文件 | 说明 | 在线预览 |
|------|------|----------|
| 研究报告PDF | 12页完整报告（含图表） | [点击在线查看](https://github.com/wdnmdlgb/Gymshark-Brand-Research/blob/master/report/Gymshark%E5%93%81%E7%89%8C%E6%B7%B1%E5%BA%A6%E7%A0%94%E7%A9%B6%E6%8A%A5%E5%91%8A.pdf) |
| 品牌能力雷达图 | 三品牌七维对比 | [交互图表](https://github.com/wdnmdlgb/Gymshark-Brand-Research/blob/master/charts/01_radar_brand_comparison.html) |
| 价格带柱状图 | 五品类价格对比 | [交互图表](https://github.com/wdnmdlgb/Gymshark-Brand-Research/blob/master/charts/02_bar_price_comparison.html) |
| 用户痛点饼图 | 负面反馈分布 | [交互图表](https://github.com/wdnmdlgb/Gymshark-Brand-Research/blob/master/charts/03_pie_user_painpoints.html) |

> 💡 PDF可直接在浏览器中预览，无需下载；HTML图表需下载后本地打开查看交互效果。

---

## 🎯 项目目标

拆解英国DTC健身服饰品牌 **Gymshark** 从车库创业到全球标杆的商业成功逻辑，横向对标 **Lululemon**、**Alphalete** 两家竞品，从品牌、产品、营销、社区、本地化、APP协同六大维度深度分析，最终总结中国健身服饰品牌出海欧洲的可复用策略与风险提示。

---

## 📁 仓库结构

```
Gymshark-Brand-Research/
├── README.md                          # 项目说明（本文件）
├── .gitignore                         # Git忽略规则
├── data/
│   └── Gymshark竞品对比分析表.xlsx     # 6个Sheet多维竞品对比表
├── report/
│   ├── Gymshark品牌深度研究报告.pdf    # 12页PDF报告（可在线预览）
│   ├── Gymshark品牌深度研究报告.pptx   # 12页PPT研究报告（可编辑）
│   ├── Gymshark品牌深度研究报告.md     # 完整报告Markdown版
│   ├── 研究报告大纲.md                 # 12页报告逐页框架
│   └── 用户反馈分析.md                 # 用户评论痛点分类深度分析
├── scripts/
│   ├── generate_excel.py              # Excel竞品对比表生成脚本
│   ├── generate_ppt.py                # PPT报告生成脚本
│   └── generate_charts_and_update_ppt.py  # 图表生成+PPT插入脚本
├── charts/                            # ECharts交互式数据可视化
│   ├── 01_radar_brand_comparison.html # 品牌多维度能力雷达图
│   ├── 02_bar_price_comparison.html   # 核心品类价格带柱状图
│   └── 03_pie_user_painpoints.html    # 用户负面反馈痛点饼图
└── materials/                         # 素材归档（图表PNG、官网截图等）
    ├── chart_radar.png
    ├── chart_bar.png
    └── chart_pie.png
```

---

## 🛠️ 使用工具与技术

| 工具 | 用途 |
|------|------|
| Excel / Pandas | 多维度竞品数据整理、对比表搭建 |
| Python (openpyxl, python-pptx) | 自动化生成Excel对比表和PPT报告 |
| ECharts | 数据可视化（雷达图、柱状图、饼图） |
| 桌面调研 | 品牌官网、公开财报、用户评论、行业报告 |

---

## 🔍 研究维度

1. **品牌基础**：创立背景、营收估值、发展历程、市场地位
2. **产品线分析**：核心爆款单品、价格带定位、上新策略
3. **营销打法**：TikTok/Instagram社媒策略、KOL分层矩阵、联盟营销
4. **社区活动**：#Gymshark66 挑战案例拆解、UGC生态
5. **欧洲本地化**：线下门店扩张、多语言站点、本地支付、学生折扣
6. **APP协同**：Training App免费引流→社区沉淀→服饰变现闭环
7. **用户口碑**：Trustpilot/Reddit/BBB用户评论痛点分类分析
8. **竞品对比**：与Lululemon、Alphalete全维度横向对标

---

## 📌 核心发现

### 品牌规模
- 2024财年营收 **£607.3m**（约7.5亿美元），连续12年增长
- 估值约 **$1.45B**，从伯明翰车库到全球DTC标杆
- 健身训练品类品牌心智份额 **36%**，超过Under Armour（35%）
- 社媒粉丝 **1800万+**，覆盖200+国家，14个本地化线上商店

### 成功五要素
1. **社区文化先行**：先建健身社区身份认同，再卖产品
2. **微KOL矩阵+联盟营销**：按效果付费，ROI可控，内容真实
3. **中端价格带卡位**：$40-80区间，避开低价白牌和高端红海
4. **免费内容生态闭环**：100%免费Training App获客→活动留存→服饰变现
5. **DTC+精选线下渐进**：线上掌握用户数据，线下旗舰店+店中店平衡成本

### 用户痛点（负面反馈分布）
| 痛点 | 占比 | 核心问题 |
|------|------|----------|
| 尺码不一致 | 30% | 各系列标准不统一，洗后缩水 |
| 面料耐用性 | 25% | 深蹲透视、缝线开裂、起球 |
| 客服售后 | 20% | 退款繁琐、响应慢 |
| 限量发售/APP | 15% | drop秒空、发售时卡顿 |
| 运动内衣支撑 | 10% | 大胸支撑不足 |

### 欧洲本地化进展
- 2025.4 阿姆斯特丹开设**欧盟首家旗舰店**
- 2026.2 进入德国：Engelhorn（曼海姆）+ Breuninger（斯图加特）店中店
- 支持Klarna分期付款、iDeal荷兰本地支付、UNiDAYS学生折扣

---

## 📊 数据可视化

项目包含三张ECharts交互图表（`charts/`目录，浏览器直接打开HTML即可交互查看）：
1. **品牌多维度能力雷达图**：Gymshark vs Lululemon vs Alphalete 七维能力对比
2. **核心品类价格带柱状图**：五品类三品牌价格横向对比
3. **用户负面反馈饼图**：五大痛点占比分布

同时在 `materials/` 目录提供对应PNG图片，已嵌入PPT报告。

---

## 📎 产出文件说明

| 文件 | 说明 |
|------|------|
| `report/Gymshark品牌深度研究报告.pdf` | **12页PDF报告，GitHub可直接在线预览**，含3张数据可视化图表 |
| `report/Gymshark品牌深度研究报告.pptx` | 12页可编辑PPT，深蓝商务配色 |
| `report/Gymshark品牌深度研究报告.md` | 完整12页报告文字版，含图表分析嵌入 |
| `report/用户反馈分析.md` | 基于公开平台用户评论的5类痛点深度拆解 |
| `data/Gymshark竞品对比分析表.xlsx` | 6个Sheet：品牌基础/产品线价格/营销打法/欧洲本地化/APP对比/爆款单品 |
| `scripts/generate_excel.py` | 可复用的Excel生成脚本，运行后自动生成6-Sheet对比表 |
| `scripts/generate_ppt.py` | 可复用的PPT生成脚本，运行后自动生成12页报告 |
| `scripts/generate_charts_and_update_ppt.py` | 图表生成+PPT插入脚本 |
| `scripts/ppt_to_pdf.py` | PPT转PDF脚本 |

---

## 📚 数据来源

- Gymshark官网及14个本地化站点
- Companies House公开财报（FY2024）
- App Store / Google Play用户评价
- Trustpilot / Sitejabber / BBB用户投诉
- Morning Consult品牌心智份额报告（2026）
- Modash.io KOL营销数据分析
- Ben Francis个人博客（benfrancis.com）
- 行业媒体：Business Live、SGB Media、Hugo跨境等
- Reddit r/Gymshark社区讨论

> ⚠️ 本项目全部基于公开网络信息做桌面研究，未获取企业内部私有数据；仅用于个人学习与求职作品集，与Gymshark公司无任何关联。

---

## 🔗 关联项目

本项目是个人求职作品集三个项目之一：
- **项目1**：Gymshark品牌深度研究（本仓库，市场/产品/BD方向）
- **项目2**：健身APP用户行为数据分析（Python+SQL+PowerBI，数据方向）
- **项目3**：欧洲居家健身Ins账号模拟运营（海外社媒运营方向）

---

*Made with ❤️ for portfolio | 2026*
