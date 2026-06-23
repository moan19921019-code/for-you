# 合同审核工具 v3

**规则引擎 + AI增强 + 民法典背书** — 三层审核体系

---

## 三级部署

| 级别 | 方式 | 需要 | 能力 |
|---|---|---|---|
| **L0 · 零门槛** | 双击 `合同审核工具.html` | 浏览器 | DOCX + AI审核 + 民法典 |
| **L1 · 分享版** | GitHub Pages | push 代码 | 同上，URL 可访问 |
| **L2 · 完整版** | L0 + 启动后端 | Python + 依赖 | + DOC / PDF / OCR + 后端规则引擎 |

## 快速开始

### L0 模式（推荐）
1. 双击 `合同审核工具.html`
2. 点击右上角 🔑 **API** 配置 DeepSeek API Key
3. 上传 DOCX 合同即可审核

### L2 模式（需要后端）
```bash
# 安装依赖
pip install flask flask-cors python-docx PyMuPDF

# 启动后端
双击 启动合同审核.bat
# 或: python contract_audit_service.py
```

## 三层审核流水线

```
合同上传
  │
  ├── L2: 后端多格式解析 (DOCX/DOC/PDF + OCR)
  └── L0: 前端 Mammoth.js DOCX 解析
  │
  ▼
【第一层】规则引擎 · 0.1秒 · 确定性
  合同编号 / 甲乙双方 / 银行账户 / 签署日期 / 条款存在性
  │
  ▼
【第二层】AI 语义审核 · 3-8秒 · DeepSeek
  条款合理性 / 风险识别 / 缺失条款 / 措辞建议
  │
  ▼
【第三层】民法典背书 · 本地 <0.1秒
  关键词召回 → AI逐条核对 → 展示法条依据
```

## 核心功能

### 🔍 规则引擎
- 正则表达式精确匹配
- 文本包含检测
- 必须存在性检查
- 每条规则可独立启用/禁用

### 🤖 AI 增强
- DeepSeek API 语义分析
- 条款风险自动识别
- 修改建议生成
- 段落 AI 优化（Quill.js 富文本编辑）

### ⚖️ 民法典集成
- 568KB 结构化数据
- 本地关键词召回 + AI 核对
- 独立民法典查询页 (`civil-code.html`)

### ✏️ 规则编辑器
- 可视化编辑审核规则
- 实时测试匹配
- 导入/导出 JSON
- 本地存储 + 后端同步

### 💬 AI 合同助手
- 浮动聊天窗口
- 自动关联合同上下文
- 民法典条文自动查询

### 🎨 设计
- Geist 字体 + Bento Grid 布局
- 骨架屏加载态
- Pass/Warn/Error 三态转换
- 服务在线/离线指示灯

## 快捷键

| 快捷键 | 功能 |
|---|---|
| `Ctrl+R` | 打开规则编辑器 |
| `Esc` | 关闭面板 |

## 文件结构

```
合同审核/
├── index.html                 # 主页面
├── civil-code.html            # 民法典查询页
├── contract_audit_service.py  # Flask 后端
├── audit_rules.json           # 审核规则配置
├── 启动合同审核.bat            # 后端启动脚本
├── css/
│   └── quill.snow.css
├── js/
│   ├── civil-code-data.js     # 民法典全文
│   ├── civil-code-evidence.js # 民法典匹配引擎
│   ├── civil-code-browser.js  # 民法典浏览器
│   ├── civil-code-page.js     # 民法典页面逻辑
│   ├── mammoth.browser.min.js # DOCX 解析
│   ├── jszip.min.js           # DOCX 格式保留
│   └── quill.min.js           # 富文本编辑
└── pics/                      # 截图
```

## API 端点（L2 模式）

| 方法 | 路径 | 说明 |
|---|---|---|
| POST | `/audit` | 上传文件 → 提取文本 + 规则审核 + 条款检测 |
| POST | `/extract` | 仅提取文本（供前端 AI 使用） |
| GET | `/rules` | 获取服务端规则配置 |
| POST | `/rules` | 保存规则配置到服务端 |
| GET | `/health` | 健康检查 |
