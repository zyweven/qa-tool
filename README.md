# qa-tool

把 xlsx 当数据库用的 Q&A 管理页面。左边看列表,右边加新条目。

## 适用场景

- 维护 Q&A 知识库(面试准备 / 培训资料 / FAQ)
- 不想每次开 Excel 操作
- 需要多人/多设备维护同一份表

## 环境要求

- **Python 3.10+**(`str | None` 这种语法需要 3.10+)
- **macOS** 推荐(用了 macOS 原生文件选择框);Linux/Windows 也行,只是文件选择需要手动输入路径

## 安装与启动

```bash
# 1. 把 qa-tool 整个文件夹放到任何位置(比如 ~/qa-tool)
# 2. 进入目录
cd ~/qa-tool

# 3. 建虚拟环境 + 装依赖(首次)
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 4. 启动(指定 xlsx 文件)
python qa_web.py /path/to/你的问答.xlsx

# 5. 浏览器打开 http://localhost:5000
```

> 第二次启动只需要第 4、5 步(venv 已建好)。

### 不传参数启动

macOS 上不传 xlsx 路径,会弹原生文件选择框。

## 功能速览

| 操作 | 怎么用 |
|---|---|
| 添加 Q&A | 右侧表单 → Cmd+Enter 提交 |
| 查看详情 | 点左侧列表项 → 弹模态框(可复制问题/答案) |
| 搜索 | 左侧搜索框(同时匹配问题和答案) |
| 切换文件 | 顶部"切换文件"按钮 |
| 合并 xlsx | 顶部"合并文件"按钮 → 选源文件 → 预览 → 合并 |
| **新建 xlsx** | 顶部"新建文件"按钮 → 选保存位置 → 创建(自动切换) |
| 关闭模态框 | ESC 键 / 点灰色背景 / 关闭按钮 |

## 答案字符限制

- 上限 **2000 字符**
- 按 **Unicode code point** 计数(中文=1、emoji=1、空格=1、英文=1)
- 超出时输入框变红 + 添加按钮禁用
- 前端 + 后端双重校验,绕过前端也会被后端拒绝

## 合并行为

- 源文件 Q&A 追加到当前文件
- 相同问题(按问题文字精确匹配)自动**跳过**(可恢复,合并前会自动备份)
- 合并后所有序号**重制为 1, 2, 3, ... N**
- 合并前自动备份:`<原文件名>.backup_YYYYMMDD_HHMMSS`

## 文件格式要求

xlsx 必须至少有 3 列,表头要包含:
- **序号** 列(支持别名:编号 / ID / Seq / No / #)
- **问题** 列(支持别名:Question / Q)
- **答案** 列(支持别名:回答 / Answer / A)

非 Q&A 格式的文件(财务报表、客户名单等)在打开/合并时会被拒绝并报明确错误。

## 数据备份建议

工具本身**不会**自动备份(除合并前的临时备份)。
建议:
- 关键修改前手动备份(把 xlsx 复制一份)
- 用 Git 跟踪 xlsx(虽然 xlsx 是二进制,但能看 diff 和回滚)

## 常见问题

**Q: 端口 5000 被占用?**
A: 改 `qa_web.py` 最后一行的 `port=5000`,比如改成 5001。

**Q: 别的电脑怎么访问这个服务?**
A: 把 `qa_web.py` 里的 `host="127.0.0.1"` 改成 `host="0.0.0.0"`,然后用 `http://<本机IP>:5000` 访问(注意防火墙)。

**Q: openpyxl 报 `No module named 'openpyxl.packaging.extended'`?**
A: 重建 venv:`rm -rf venv && python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt`

**Q: 怎么把同一份数据同步到另一台电脑?**
A: 把 xlsx 文件拷过去(用 iCloud / U盘 / Git 都行),在另一台电脑上启动时指定这个 xlsx 即可。

## 项目结构

```
qa-tool/
├── qa_web.py            # Flask 后端
├── templates/
│   └── index.html       # 前端(单文件)
├── requirements.txt     # flask + openpyxl
├── test_api.py          # 端到端测试(可选)
├── README.md            # 本文件
└── venv/                # 虚拟环境(本地建,不要提交)
```

## API 速查(供二次开发)

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/api/info` | 当前 xlsx 路径 + 状态 |
| GET | `/api/qa?q=关键词` | 列表/搜索 |
| POST | `/api/qa` | 新增 `{question, answer, force_overwrite?}` |
| POST | `/api/xlsx` | 切换文件 `{path}` |
| POST | `/api/merge` | 合并 `{source_path}`(先备份再合并并重排序号) |
| POST | `/api/preview-merge` | 预览合并(不实际写) |
| POST | `/api/pick-file` | 弹 macOS 原生文件选择框 |
