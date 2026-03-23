# 安装和配置指南

为新的使用者准备的完整设置步骤。

---

## 系统要求

- Python 3.7 及以上
- Windows / Mac / Linux
- 500 MB 硬盘空间（用于依赖包）

---

## Step 1: 安装 Python 包

```bash
pip install openpyxl pandas python-docx pillow

# 或者用 requirements.txt（如果有）
pip install -r requirements.txt
```

### 验证安装

```bash
python -c "import openpyxl, pandas, docx, PIL; print('OK')"
```

如果输出 `OK`，说明全部安装成功。

---

## Step 2: 下载 Skill 文件

复制整个 `reward-audit` 文件夹到你的电脑：

```
reward-audit/
├── README.md                       （新用户从这里开始）
├── SKILL.md                        （完整文档）
├── QUICK_REFERENCE.md              （快速参考卡）
├── SETUP.md                        （本文件）
├── scripts/
│   ├── audit_to_word.py           （原始版，不推荐）
│   └── gen_audit_report_v2.py      （改进版，推荐用这个）
└── references/
    ├── game_economy.md            （游戏经济参考）
    └── case_library.md            （7 个真实案例）
```

---

## Step 3: 配置你的游戏经济参数

编辑 `references/game_economy.md`：

### 3.1 更新货币汇率

找到这一段，改成你游戏的数据：

```markdown
## 货币汇率参考

| 换算关系 | 比例 | 备注 |
|---------|------|------|
| 筹码 → MS | 1 筹码 = 0.43 MS | 你的游戏可能是不同比例 |
```

**怎么填**：
- 如果你的游戏用「金币」和「钻石」
  - 就改成：`1 金币 = 0.1 MS` 之类的
- 不确定就问项目经理

### 3.2 更新道具价值

找到这一段：

```markdown
## 道具价值参考

| 道具名称 | 稀有度 | 参考价值（MS） | ... |
|---------|------|---------|---------|
| 超星灵药精华 | 稀有 | 1.35 MS/个 | ... |
```

**怎么填**：
- 列出你游戏里的 5-10 个常见道具
- 标记它们的 MS 价值
- 示例：
  ```
  | 钻石 | 稀有 | 1.0 MS/个 |
  | 经验书 | 普通 | 0.5 MS/个 |
  | 限定道具 | 传说 | 10.0 MS/个 |
  ```

### 3.3 更新资源产出预期

找到这一段：

```markdown
## 各阶段资源产出预期

| 游戏阶段 | 日均筹码产出 | 日均精华产出 | 说明 |
|---------|------------|-----------|------|
| 新手期 | 100-200 | 0-2 | 低等级玩家参与度低 |
```

**怎么填**：
- 根据你游戏的数据填入每个阶段的产出量
- 不确定可以估算：活动期总产出 ÷ 天数

---

## Step 4: 修改脚本配置

编辑 `scripts/gen_audit_report_v2.py`：

### 4.1 改 Excel 文件路径

找到这一行（大约第 18 行）：

```python
XLSX = r'C:\Users\Administrator\AppData\Roaming\im\628670@nd\RecvFile\熊雨微_392133\【魔域】25年冰雪派对 活动奖励案（非怀旧）v1.0+.xlsx'
```

改成你要审核的文件：

```python
XLSX = r'D:\Game_Data\【你的游戏名】活动奖励.xlsx'
```

### 4.2 改输出路径

找到这一行（大约第 20 行）：

```python
REPORT_PATH = r'E:\桌面\奖励审核报告_冰雪派对_带截图版_V2.docx'
```

改成你想保存报告的位置：

```python
REPORT_PATH = r'D:\Reports\奖励审核报告_2026-03-24.docx'
```

### 4.3 改检查的工作表和问题

找到这一段（大约第 40 行开始的 `issues = [`）：

**原样本**：

```python
issues = [
    {
        'sheet': '1在线+每日参与+活跃礼包',
        'rows': [9, 13],
        'cols': [1, 18],
        'title': '在线/活跃礼包价值列全为0',
        ...
    },
    ...
]
```

**怎么改**：

1. 打开你的 Excel 文件，看工作表名称
2. 在 `issues` 列表里添加对应的检查项
3. `'sheet'`：工作表名
4. `'rows'`：要检查的行范围（1-based）
5. `'cols'`：要检查的列范围（1-based）
6. `'title'`：问题标题（简短说明）

**例子**：

```python
issues = [
    {
        'sheet': '日常任务',
        'rows': [2, 20],           # 检查第 2-20 行
        'cols': [1, 5],            # 检查第 A-E 列
        'title': '任务奖励价值统计缺失',
        'severity': '[建议]',
        'desc': '价值列没有填写',
        'impact': '无法评估总成本',
        'suggestion': '填入计算后的 MS 价值'
    }
]
```

---

## Step 5: 测试运行

确保一切配置正确：

```bash
python scripts/gen_audit_report_v2.py
```

### 预期输出

```
[1/4] 初始化...
[2/4] 生成表格截图...
  生成图片: 日常任务 R2:R20
  -> C:\...\issue_1.png
  ...
[3/4] 生成 Word 报告...
[OK] 报告已保存：D:\Reports\奖励审核报告_2026-03-24.docx
[4/4] 完成！
验证成功：42.5 KB
```

### 常见问题排查

| 问题 | 原因 | 解决 |
|------|------|------|
| `ModuleNotFoundError: No module named 'openpyxl'` | 没装依赖 | 运行 `pip install openpyxl ...` |
| `FileNotFoundError: No such file or directory: '...'` | 路径错了 | 检查 XLSX 路径是否存在 |
| `PermissionError` | Excel 文件被锁定 | 关闭 Excel，或改个输出路径 |
| 报告没有图片 | 脚本版本错了 | 确保用的是 `gen_audit_report_v2.py` |
| 卡住了 | 表格太大或系统卡 | 等待或重启试试 |

---

## Step 6: 使用 WorkBuddy 加载 Skill（可选）

如果你用 WorkBuddy，可以直接加载这个 skill：

1. 在 WorkBuddy 里选择「加载 Skill」
2. 选择 `reward-audit` 文件夹
3. 告诉 AI：「审核这个 Excel 奖励文件」
4. AI 会自动调用脚本

---

## 后续运维

### 定期更新参考数据

```
每个季度检查一次：
- [ ] 游戏的汇率是否变了（game_economy.md）
- [ ] 有无新的典型错误（case_library.md）
- [ ] 脚本的检查项是否还适用
```

### 积累新案例

```
每发现一个新错误，就加到 case_library.md：

### 案例编号：CASE-008
- **问题类型**：（填入）
- **发生场景**：（填入）
- **错误表现**：（填入）
...
```

---

## 常见定制需求

### 需求 1：改变报告格式

编辑脚本中的「生成 Word 报告」部分（大约第 120 行）：

```python
doc.add_heading(f'问题{idx+1}: {issue["title"]}', level=1)  # 改 heading 级别
doc.add_paragraph(...)  # 添加自定义段落
```

### 需求 2：添加自动化检查

在 `issues` 前面加代码，自动分析 Excel：

```python
import openpyxl
wb = openpyxl.load_workbook(XLSX, data_only=True)

# 自动找出所有价值为 0 的单元格
issues = []
for sheet in wb.sheetnames:
    ws = wb[sheet]
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == 0:  # 找到 0 值
                issues.append({
                    'sheet': sheet,
                    'title': f'发现零值...',
                    ...
                })
```

### 需求 3：生成不同格式的报告

- **Excel 报告**：用 `openpyxl` 写回 Excel
- **PDF 报告**：用 `python-docx` 转 PDF（需要额外库）
- **HTML 报告**：用模板引擎生成 HTML

---

## 常见问题（FAQ）

### Q: 我的 Excel 在 Mac，能用吗？
**A:** 能。脚本跨平台支持，注意路径分隔符改成 `/`。

### Q: 能在服务器上自动运行吗？
**A:** 能。用 `cron`（Mac/Linux）或 `Task Scheduler`（Windows）定时运行。

### Q: 支持批量处理多个文件吗？
**A:** 目前不支持，但可以手动改脚本循环处理。

### Q: 报告能直接发给老板吗？
**A:** 能。生成的是标准 Word 文档，可以直接分享。

---

## 获得帮助

| 问题 | 查看 |
|------|------|
| 脚本运行出错 | 本文件的「常见问题排查」 |
| 不知道怎么用 | README.md |
| 不知道在找什么问题 | case_library.md （7 个真实案例） |
| 脚本如何工作 | SKILL.md |
| 快速看概览 | QUICK_REFERENCE.md |

---

**祝使用愉快！** 🎉

版本 1.0 | 2026-03-23
