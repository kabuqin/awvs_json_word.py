# awvs_json_word.py 使用说明手册

Acunetix/Invicti 漏洞扫描 JSON 报告一键生成 Word 文档工具

---

## 一、脚本简介

`awvs_json_word.py` 是一个将 **Acunetix（AWVS）/ Invicti** 扫描器导出的 JSON 结果文件，自动生成为**中文 Word 安全漏洞报告**的一体化工具。

脚本内置了数据解析、格式转换、中文化翻译、统计图表生成和报告排版等全部功能，只需一条命令即可完成从原始扫描数据到正式报告的全过程，无需任何人工干预。

### 主要功能

| 功能 | 说明 |
|------|------|
| 一键生成 | 直接输入扫描器导出的 JSON，一步输出 Word 报告 |
| 多目标支持 | 自动识别多个扫描站点，生成站点漏洞统计表 |
| 全中文报告 | 漏洞名称、漏洞描述、修复建议自动翻译为中文 |
| 统计图表 | 自动生成漏洞严重程度环形图、柱状图 |
| 风险着色 | 按严重/高危/中危/低危/信息分级着色显示 |
| 检测详情 | 自动解压还原扫描器的 HTTP 请求/响应交互数据 |
| 修复建议 | 每个漏洞附带中文修复建议，另有系统性修复汇总 |

---

## 二、环境要求

| 依赖项 | 要求 |
|--------|------|
| Python | 3.8 及以上版本 |
| python-docx | Word 文档生成库 |
| matplotlib | 统计图表绘制库（需支持中文字体） |

### 安装依赖

```bash
pip install python-docx matplotlib
```

> **Windows 用户注意**：图表中文渲染依赖 `SimHei`（黑体）字体，Windows 系统默认自带（`C:\Windows\Fonts\simhei.ttf`）。若图表中文显示为方块，请清除 matplotlib 字体缓存后重试：
>
> ```bash
> python -c "import matplotlib; print(matplotlib.get_cachedir())"
> # 删除缓存目录下的 fontlist-*.json 文件后重新运行脚本
> ```

---

## 三、使用方法

### 命令行格式

```bash
python awvs_json_word.py <report.json> [output.docx]
```

| 参数 | 必填 | 说明 |
|------|------|------|
| `report.json` | 是 | Acunetix/Invicti 导出的扫描结果 JSON 文件路径 |
| `output.docx` | 否 | 输出 Word 文件名称；省略时自动生成 `report_报告编号_时间戳.docx` |

### 使用示例

```bash
# 基本用法（自动命名输出文件）
python awvs_json_word.py scan_result.json

# 指定输出文件名
python awvs_json_word.py scan_result.json 安全漏洞扫描报告.docx

# 查看帮助
python awvs_json_word.py -h
```

### 输出位置

生成的 Word 报告与统计图表统一输出到脚本运行目录下的：

```
workspace/
└── report_result/
    ├── 报告名称.docx      # 生成的 Word 报告
    └── vuln_chart.png     # 漏洞统计图表（可单独使用）
```

---

## 四、输入文件格式说明

脚本接受 **Acunetix（AWVS）/ Invicti 扫描器**的标准 JSON 导出格式，顶层结构如下：

```json
{
  "export": {
    "lang": "cn",
    "scans": [
      {
        "info": {
          "host": "https://example.com",
          "start_url": "https://example.com/",
          "start_date": "2026-08-13T00:59:35+00:00",
          "end_date": "2026-08-13T01:04:55+00:00",
          "build": "25.1.250204093"
        },
        "vulnerability_types": [
          {
            "vt_id": "34a6c791-...",
            "name": "HTTP Strict Transport Security (HSTS) Policy Not Enabled",
            "severity": 2,
            "description": "...",
            "recommendation": "...",
            "type": "configuration"
          }
        ],
        "vulnerabilities": [
          {
            "info": {
              "vt_id": "34a6c791-...",
              "name": "HTTP Strict Transport Security (HSTS) Policy Not Enabled",
              "url": "https://example.com/",
              "details": "URLs where HSTS is not enabled: ...",
              "request": "GET / HTTP/1.1\r\nHost: example.com\r\n\r\n",
              "tags": ["confidence.100", "verified"]
            },
            "response": "H4sIAAAAAAAA..." 
          }
        ]
      }
    ]
  }
}
```

### 关键字段说明

| 字段 | 说明 |
|------|------|
| `scans` | 扫描任务列表，每个任务对应一个站点；多个站点自动生成对比统计 |
| `info.host` / `info.start_url` | 站点地址，作为站点统计表的标识 |
| `info.build` | 扫描器版本号，自动写入报告"工具版本"栏 |
| `vulnerability_types` | 漏洞类型定义，含名称、等级、描述、修复建议 |
| `vulnerabilities` | 具体漏洞实例，通过 `vt_id` 关联类型定义 |
| `vulnerabilities[].info.request` | HTTP 请求原文，自动解析为结构化交互记录 |
| `vulnerabilities[].response` | HTTP 响应（gzip 压缩的 base64），自动解压还原 |
| `vulnerabilities[].info.details` | 扫描器检测详情，当无请求/响应数据时（如 SSL 审计类漏洞）自动补充到漏洞详情中 |
| `tags` | 置信度标签，`confidence.100` 自动映射为 `verified`（已确认） |

> 若输入的 JSON 不是 Acunetix/Invicti 导出格式（缺少 `export.scans` 字段），脚本会报错并提示。

---

## 五、生成报告内容结构

生成的 Word 报告包含以下章节：

```
┌─────────────────────────────────────────────┐
│  封面页                                      │
│   - 报告标题（Web 应用安全漏洞扫描报告）      │
│   - 测试目标 / 测试时间 / 测试工具 / 漏洞总数 │
├─────────────────────────────────────────────┤
│  目  录                                      │
├─────────────────────────────────────────────┤
│  一、漏洞扫描概述                            │
│     扫描时间、工具版本、漏洞数量统计         │
├─────────────────────────────────────────────┤
│  二、漏洞风险统计                            │
│     环形图 + 柱状图 + 漏洞类型分布表         │
├─────────────────────────────────────────────┤
│  三、网站漏洞列表（多目标时）                │
│     按站点统计各等级漏洞数量                 │
├─────────────────────────────────────────────┤
│  四、漏洞详情                                │
│     每个漏洞包含：                           │
│      - 标题（含风险等级，如【中危】）        │
│      - 漏洞编号/风险等级/类型/目标地址       │
│      - 漏洞描述（中文）                     │
│      - 漏洞详情（HTTP 请求/响应交互）        │
│      - 修复建议（中文）                     │
├─────────────────────────────────────────────┤
│  五、修复建议汇总                            │
│     SQL注入/认证/信息泄露/CSRF/响应头/代码   │
└─────────────────────────────────────────────┘
```

### 风险等级对照

| 扫描器数值 | 报告等级 | 显示样式 |
|-----------|---------|---------|
| 4 | 严重（Critical） | 红色 |
| 3 | 高危（High） | 橙色 |
| 2 | 中危（Medium） | 黄色 |
| 1 | 低危（Low） | 绿色 |
| 0 | 信息（Info） | 蓝色 |

---

## 六、中文化机制

脚本内置 `VULN_ZH_MAP` 翻译表，将 Acunetix 常见漏洞的英文名称、描述、修复建议自动翻译为中文。

### 已内置翻译的漏洞类型（14 种）

| 英文名称 | 中文名称 |
|---------|---------|
| HTTP Strict Transport Security (HSTS) Policy Not Enabled | 未启用 HTTP 严格传输安全（HSTS）策略 |
| Misconfigured Access-Control-Allow-Origin Header | CORS 跨域配置错误 |
| TLS/SSL Sweet32 attack | TLS/SSL Sweet32 攻击漏洞 |
| TLS/SSL Weak Cipher Suites | TLS/SSL 弱加密套件 |
| Cookies Not Marked as Secure | Cookie 未设置 Secure 标志 |
| Cookies with missing, inconsistent or contradictory properties | Cookie 属性缺失、不一致或相互矛盾 |
| Content Security Policy (CSP) Not Implemented | 未实施内容安全策略（CSP） |
| Permissions-Policy header not implemented | 未实施 Permissions-Policy 响应头 |
| Web Application Firewall Detected | 检测到 Web 应用防火墙（WAF） |
| Insecure HTTP Usage | 使用不安全的 HTTP 协议 |
| SSL/TLS Not Implemented | 未启用 SSL/TLS 加密 |
| Access-Control-Allow-Origin header with wildcard (*) value | Access-Control-Allow-Origin 头使用通配符（*） |
| Subresource Integrity (SRI) Not Implemented | 未实施子资源完整性（SRI）校验 |
| [Possible] Internal IP Address Disclosure | （疑似）内网 IP 地址泄露 |

### 扩展翻译表

遇到未收录的新漏洞类型时，报告会回退显示原始英文内容。如需补充翻译，编辑脚本中的 `VULN_ZH_MAP` 字典，按以下格式添加：

```python
VULN_ZH_MAP = {
    "漏洞英文名称（与扫描器导出一致）": {
        "title_zh": "漏洞中文名称",
        "description_zh": "漏洞中文描述",
        "repair_zh": "漏洞中文修复建议",
    },
    # ... 其他漏洞
}
```

---

## 七、注意事项

1. **输入校验**：脚本会校验 JSON 是否为 `export.scans` 结构，格式不符会立即报错退出。
2. **输出文件占用**：若同名输出文件正被 Word 打开，脚本会因文件占用报错，请关闭后重试或更换输出文件名。
3. **图表字体**：matplotlib 首次运行会构建字体缓存，属正常现象；中文显示异常时按"二、环境要求"中的方法清除缓存。
4. **缺失数据容错**：漏洞无请求/响应数据时，脚本自动改用扫描器的 `details` 检测详情填充"漏洞详情"章节，不会出现空白漏洞。
5. **响应截断**：超长 HTTP 响应体（超过 1500 字符）会被截断并标注"内容已截断"，避免报告篇幅失控。
6. **报告编号**：封面与页脚不显示报告编号，仅内部使用（用于自动命名输出文件）。

---

## 八、常见问题（FAQ）

**Q1：运行报错 `ModuleNotFoundError: No module named 'docx'`**
依赖未安装，执行 `pip install python-docx matplotlib`。

**Q2：报告图表中文显示为方块**
SimHei 字体未被 matplotlib 缓存识别，删除 `%LOCALAPPDATA%\matplotlib` 下的 `fontlist-*.json` 后重新运行。

**Q3：漏洞标题/描述显示英文**
该漏洞类型未收录在 `VULN_ZH_MAP` 翻译表中，按"六、中文化机制"中的方法自行补充翻译。

**Q4：为什么有些漏洞的"漏洞详情"只有响应没有请求**
SSL 审计类漏洞（如 Sweet32、弱加密套件）本身无 HTTP 请求数据，脚本自动使用扫描器的检测详情（details 字段）填充，属正常现象。

**Q5：如何同时生成多个扫描任务的报告**
Acunetix 导出的 JSON 中若包含多个 `scans`（多个站点），脚本会自动合并处理，生成包含站点对比统计的多目标报告。

---

## 九、版本信息

| 项目 | 内容 |
|------|------|
| 脚本名称 | awvs_json_word.py |
| 测试工具 | Acunetix（版本自动读取扫描数据中的 build 字段） |
| 输出格式 | Microsoft Word（.docx） |
| 更新日期 | 2026-08-16 |

---

*本手册随脚本功能同步更新，如有疑问请联系工具维护者。*
