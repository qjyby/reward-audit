---
name: reward-audit
description: 这个 skill 专为游戏数值策划设计，用于审核任务/活动奖励配置的合理性与逻辑正确性。当用户需要检查奖励数据、分析奖励结构、发现奖励逻辑错误、或验证奖励计算公式时，应加载本 skill。支持传入 xlsx 文件，自动逐工作表审核，将错误区域截图并生成带图文说明的 Word 审核报告。
---

# 任务活动奖励审核 Skill

## 职责定位

本 skill 扮演一位经验丰富的游戏数值策划助手，专注于**任务/活动奖励配置的审核工作**。

核心职责：
- 检查奖励数值的合理性（是否符合游戏整体经济体系）
- 识别奖励配置中的逻辑性错误（计算错误、条件错误、奖励缺失/重复等）
- 发现奖励结构设计上的潜在风险（通货膨胀、养号风险、付费破坏等）
- 给出改进建议，并说明理由
- **支持 xlsx 文件输入，自动逐工作表审核，生成带截图的 Word 报告**

---

## 审核流程（SOP）

### 第一步：理解奖励背景

在开始审核前，先确认以下信息（如用户未提供，主动询问）：
- 这是什么类型的任务/活动？（日常任务、周任务、限时活动、成就系统……）
- 目标玩家群体是谁？（新手、中期、高战力、全体……）
- 奖励周期是多长？（单次、每日、每周、活动期间……）
- 这份奖励在游戏中处于什么阶段？（开服初期、中期成熟期、活动运营期……）

### 第二步：解析奖励结构

拿到奖励数据后，按以下维度拆解：

1. **奖励类型分类**
   - 货币类（金币、钻石、代币……）
   - 道具类（材料、装备、消耗品……）
   - 经验类（角色经验、装备经验……）
   - 稀有资源类（限定道具、高价值材料……）

2. **数量级核查**
   - 与同类型任务/活动的奖励进行横向对比
   - 计算单位时间/单位投入的产出比
   - 判断是否存在数量级异常（多一个零或少一个零）

3. **奖励梯度核查**
   - 高难度/高门槛是否对应更高奖励
   - 奖励曲线是否平滑（有无断层或倒挂）
   - 首通、多次完成的奖励差异是否合理

### 第三步：逻辑错误检查

重点排查以下常见错误：

| 错误类型 | 说明 | 检查方式 |
|---------|------|---------|
| 计算错误 | 总奖励 ≠ 各子项加总 | 逐项加总验证 |
| 条件逻辑错误 | 触发条件与奖励不匹配 | 逐条件对照奖励配置 |
| 奖励缺失 | 某个阶段/档位没有奖励 | 检查覆盖率是否完整 |
| 奖励重复 | 同一目标被多次奖励 | 查找重复的奖励来源 |
| 上下限缺失 | 可重复领取奖励没有每日/总上限 | 检查 cap 设置 |
| 奖励倒挂 | 低难度获得更好奖励 | 难度-奖励排序对比 |
| 汇率换算错误 | 不同货币换算比例不合理 | 参照游戏内汇率体系 |

### 第四步：生成 Word 审核报告（含错误截图）

当用户传入 xlsx 文件时，使用 `scripts/audit_to_word.py` 脚本自动化完成以下工作：

1. **读取 xlsx 所有工作表**，逐表进行数据审核分析
2. **标记错误单元格**：用红色高亮标注有问题的行/列区域
3. **截图保存**：将每处错误区域截图为 PNG 图片
4. **生成 Word 文档**：每个问题在 Word 里以「截图 + 文字描述」成对呈现

#### 脚本调用方式

```bash
python scripts/audit_to_word.py <xlsx文件路径> [输出目录]
```

- `xlsx文件路径`：用户传入的奖励配置 Excel 文件
- `输出目录`（可选）：Word 报告和截图的保存位置，默认与 xlsx 同目录

#### Word 报告结构

```
奖励审核报告 - [文件名]
审核时间：YYYY-MM-DD HH:MM

一、总体概况
  - 审核工作表数量：X 张
  - 发现问题总数：X 个（严重 X / 中等 X / 建议 X）

二、逐工作表审核结果

  【工作表：Sheet1 - 日常任务奖励】
  问题 1：[问题标题]
  严重等级：🔴 严重
  [错误区域截图]
  错误描述：第 N 行，总奖励数值与子项加总不符，配置值为 XXX，实际应为 YYY。
  修改建议：将该行总奖励修改为 YYY。

  问题 2：...

  【工作表：Sheet2 - 活动奖励】
  ...

三、汇总建议
  [综合性改进方向]
```

---

## 审核原则

- **数据说话**：所有结论需基于具体数据，不做无依据的主观判断
- **系统视角**：单个奖励不孤立看待，要放入整体经济体系判断
- **玩家视角**：从玩家实际体验出发，判断奖励是否有吸引力且公平
- **风险优先**：发现严重问题（如无上限的资源产出）要优先且醒目地标注
- **建设性反馈**：指出问题的同时，给出具体可操作的修改方向
- **全表覆盖**：xlsx 中每一张工作表都必须审核，不遗漏

---

## 常用参考维度

### 奖励价值参考体系（待用户补充）
> 提示：可以在 references/game_economy.md 中补充本游戏的货币汇率、道具价值表、各阶段资源产出预期等参考数据，以提高审核精准度。

### 典型问题案例库（持续积累）
> 提示：随着使用过程中发现更多典型错误模式，可以在 references/case_library.md 中记录，用于后续审核的参照。

---

## 实施经验与最佳实践

### 生成 Word 报告的性能优化

**问题背景**：第一版脚本用 matplotlib 渲染大型 Excel 表格容易卡住、耗时长。

**解决方案**：改用 PIL (Pillow) 直接生成表格图片
- matplotlib 渲染完整表格：耗时长、容易卡住（matplotlib 需要绘制所有图形元素）
- PIL 文本 + 表格格线：快 100 倍，稳定性更好
- 代码示例见 `scripts/gen_audit_report_v2.py`

```python
# PIL 方式（推荐）
from PIL import Image, ImageDraw, ImageFont

img = Image.new('RGB', (img_width, img_height), color=(255, 255, 255))
draw = ImageDraw.Draw(img)
# 逐行逐列绘制数据和格线
for i, row_data in enumerate(table_data):
    for j, cell_text in enumerate(row_data):
        draw.rectangle([x, y, x + cell_width, y + cell_height], outline=(0,0,0), width=1)
        draw.text((x+5, y+5), text, fill=(0,0,0), font=font)
img.save(output_path, 'PNG')
```

### 工作流中的进度提示和错误处理

**关键改进**：
1. **分步骤进度显示**：`[1/4] 初始化...` → `[2/4] 生成截图...` → `[3/4] 生成报告...` → `[4/4] 完成！`
2. **单独的错误捕获**：每个问题都 try-catch，一个出错不影响其他
3. **中间反馈**：每完成一个问题都 print，让用户看到实时进度
4. **生成后验证**：完成后立即检查文件是否真的存在和大小是否合理

```python
print("[1/4] 初始化...")
# 初始化工作...

print("[2/4] 生成表格截图...")
for idx, issue in enumerate(issues):
    try:
        # 生成截图
        issue['screenshot'] = screenshot_path
        print(f"  问题{idx+1}/{len(issues)}: {issue['title']} [OK]")
    except Exception as e:
        print(f"  问题{idx+1} 出错: {str(e)}")
        issue['screenshot'] = None  # 标记失败但继续

print("[3/4] 生成 Word 报告...")
# 生成报告...
doc.save(REPORT_PATH)
print(f"[OK] 报告已保存：{REPORT_PATH}")

# 立即验证
if os.path.exists(REPORT_PATH):
    file_size = os.path.getsize(REPORT_PATH) / 1024
    print(f"验证成功：{file_size:.1f} KB")
```

### 编码问题处理（Windows + 中文）

**常见问题**：
- PowerShell 默认使用 GBK/GB2312 编码
- Python 代码中的 emoji 和 Unicode 字符会报 `UnicodeEncodeError`

**解决方案**：
1. **避免 emoji**：用 `[OK]` / `[ERROR]` 纯文字替代 ✓ ✗
2. **脚本编码声明**：`# -*- coding: utf-8 -*-` 在文件开头
3. **环境变量**（如需要）：`$env:PYTHONIOENCODING = 'utf-8'`
4. **不依赖 print 的返回值**：关键结果写到文件，不用 print

### 文件路径处理

**建议**：
- 用原始字符串：`r'C:\path\to\file'`（避免转义符问题）
- 或用正斜杠：`'C:/path/to/file'`
- 用 `os.path.join()` 拼接路径，不要手动拼接字符串
- 生成后用 `Test-Path` 或 `os.path.exists()` 验证

### 报告文件验证清单

生成报告后的检查清单：

```python
from docx import Document

doc = Document(report_path)

# 检查内容
print(f'段落数: {len(doc.paragraphs)}')
print(f'表格数: {len(doc.tables)}')

# 检查图片数量
image_count = 0
for rel in doc.part.rels.values():
    if 'image' in rel.reltype:
        image_count += 1

print(f'图片数: {image_count}')

# 如果预期 7 张图片，image_count 应该等于 7
assert image_count == 7, f"期望 7 张图片，实际 {image_count} 张"
```
