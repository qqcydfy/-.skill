# 周报生成技能 (weekly-report)

为仿真分析类项目生成专业周报，输出为Word文档格式。生成的周报内容可直接复制到最终技术报告中使用。

## 功能特点

- 支持多种仿真项目类型（热学、光学、结构、电磁、流体等）
- 自动生成Word文档，格式规范
- 支持自动插入图片（按文件名或序号匹配）
- 内容组织专业，可直接用于技术报告
- 支持自定义模板

## 目录结构

```
weekly-report/
├── SKILL.md                    # 技能说明文件
├── README.md                   # 本文件
├── scripts/
│   └── generate_docx.py        # 周报生成脚本
└── references/
    └── config_example.json     # 配置文件示例
```

---

## 安装方式

### 方式一：项目级安装（推荐）

适用于当前项目，只在该项目目录下生效。

**步骤：**

1. 进入你的项目根目录
```bash
cd /path/to/your/project
```

2. 创建 skills 目录（如果不存在）
```bash
mkdir -p .claude/skills
```

3. 复制 weekly-report 文件夹到项目 skills 目录
```bash
cp -r /path/to/weekly-report .claude/skills/
```

4. 验证安装
```bash
ls .claude/skills/weekly-report/
# 应显示：SKILL.md  README.md  scripts/  references/
```

安装完成后，在该项目目录下启动 Claude Code 即可使用 `/weekly-report` 命令。

**目录结构示例：**
```
your-project/
├── .claude/
│   └── skills/
│       └── weekly-report/
│           ├── SKILL.md
│           ├── README.md
│           ├── scripts/
│           └── references/
├── src/
├── data/
└── ...
```

---

### 方式二：用户级安装

适用于所有项目，在任何目录下都可以使用。

**步骤：**

1. 进入用户 Claude 配置目录
```bash
cd ~/.claude/skills
# Windows: cd %USERPROFILE%\.claude\skills
```

2. 如果目录不存在，创建它
```bash
mkdir -p ~/.claude/skills
```

3. 复制 weekly-report 文件夹
```bash
cp -r /path/to/weekly-report ~/.claude/skills/
```

4. 验证安装
```bash
ls ~/.claude/skills/weekly-report/
# 应显示：SKILL.md  README.md  scripts/  references/
```

**目录结构示例：**
```
~/.claude/
└── skills/
    └── weekly-report/
        ├── SKILL.md
        ├── README.md
        ├── scripts/
        └── references/
```

---

### 安装方式对比

| 特性 | 项目级安装 | 用户级安装 |
|------|-----------|-----------|
| 生效范围 | 仅当前项目 | 所有项目 |
| 安装位置 | `.claude/skills/` | `~/.claude/skills/` |
| 适用场景 | 项目专属技能 | 通用技能 |
| 版本控制 | 可纳入 git | 独立于项目 |

---

### 安装后验证

在 Claude Code 中输入以下命令验证安装成功：

```
/skills
```

在弹出的技能列表中应能看到 `weekly-report`。

或者直接测试：
```
/weekly-report
```

如果提示提供周报信息，说明安装成功。

---

## 快速开始

### 1. 安装依赖

```bash
pip install python-docx
```

### 2. 调用方式

**方式一：使用斜杠命令**
```
/weekly-report
```

**方式二：自然语言触发**
- "帮我写周报"
- "生成本周工作总结"
- "整理仿真结果写周报"
- "仿真结果整理"

调用后，Claude会询问你需要的信息，按提示提供即可。

---

## 需要准备的材料

| 材料 | 必须 | 说明 |
|------|------|------|
| 时间范围 | 是 | 如：5.5-5.11 |
| 作者姓名 | 是 | 如：xxx |
| 工作内容 | 是 | 本周做了什么，分模块描述 |
| 图片目录 | 否 | 仿真结果图片所在文件夹路径 |
| 总结 | 否 | 关键发现、下周计划 |

---

## 图片文件夹指定方式

### 方式1：调用时直接说明

```
/weekly-report
图片目录是 ./pic
```

### 方式2：在对话中提供

```
帮我写周报，图片在 ./simulation_results/ 文件夹里
```

### 图片命名建议

- 按顺序命名：`fig1.png`, `fig2.png`, `fig3.png`...
- 或带关键词：`temperature_field.png`, `spectrum.png`...
- 支持格式：png, jpg, jpeg, bmp, tiff, gif

---

## 周报内容模板

向Claude提供以下信息即可生成周报：

```
时间范围：5.5-5.11
作者：xxx
项目：VCSEL芯片热学仿真

模块1：芯片布局设计
- 4颗VCSEL芯片以4×1阵列排列
- 阵列尺寸：0.98mm×0.23mm×0.2mm

模块2：温度场仿真
- 热源功率：58.8mW
- TEC热面温度：50℃
- 结果：最大温差0.012K

模块3：位置优化研究
- 1）X轴移动研究
  - X轴扫描步长：1mm
  - 结论：X轴位置对温差影响可忽略
- 2）对角线移动研究
  - 对角线扫描步长：1mm
  - 结论：中心位置冷却效果最佳

总结：
- 位置对最大温差影响可忽略
- 中心位置冷却效果更好

图片目录：./pic
```

---

## 输出格式说明

生成的Word文档结构：

```
5.5-5.11周报 xxx                    ← 22pt，居中

进展：                                  ← 18pt
一段话概述本周工作                       ← 15pt

①芯片布局设计                          ← 16pt，加粗
详细内容...                             ← 15pt

[图片]                                  ← 居中，宽15cm
图 1 芯片布局示意图                     ← 10pt，居中，黑色
备注：数据来源说明                       ← 15pt

                                        ← 空2行分隔

②温度场仿真                            ← 16pt，加粗
详细内容...

图 2 温度场分布                         ← 10pt，居中，黑色

③位置优化研究
详细内容...

3）总结                                ← 16pt，加粗
关键发现...                             ← 15pt
```

### 格式规范

| 元素 | 字号 | 加粗 | 对齐 |
|------|------|------|------|
| 周报标题 | 22pt | 否 | 居中 |
| 进展概述 | 18pt | 否 | 左对齐 |
| 模块标题（①②③） | 16pt | 是 | 左对齐 |
| 正文内容 | 15pt | 否 | 左对齐 |
| 子标题 | 15pt | 是 | 左对齐 |
| 图片说明 | 10pt | 否 | 居中 |
| 图片颜色 | - | - | 黑色 |

---

## 使用示例

### 示例1：交互式生成

```
用户：/weekly-report

Claude：请提供以下信息：
1. 时间范围
2. 作者姓名
3. 工作内容（分模块）
4. 图片目录（可选）
5. 总结（可选）

用户：
时间：5.5-5.11
作者：xxx
项目：微带天线仿真

模块1：天线单元设计
- 频率2.4GHz，FR-4基板
- 贴片尺寸28.5mm×37.2mm

模块2：S参数仿真
- S11最小值-28.6dB@2.42GHz
- 带宽3.7%

图片在 ./results/

Claude：[生成周报.docx]
```

### 示例2：命令行生成

```bash
python scripts/generate_docx.py \
  --config references/config_example.json \
  --output "周报_xxx_4.24-4.30.docx"
```

### 示例3：自动匹配图片

```bash
python scripts/generate_docx.py \
  --config config.json \
  --image-dir ./simulation_results \
  --auto-images \
  --output weekly_report.docx
```

---

## 配置文件说明

### 基本字段

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| date_range | string | 是 | 周报时间范围（如：4.24-4.30） |
| author | string | 是 | 作者姓名 |
| project | string | 否 | 项目名称 |
| summary | string | 否 | 本周工作概述 |
| image_dir | string | 否 | 图片目录路径 |
| sections | array | 是 | 章节内容数组 |
| conclusion | array | 否 | 总结内容 |

### 章节结构 (sections)

每个章节包含：

```json
{
  "title": "章节标题",
  "content": [
    "文本内容",
    "更多文本",
    {"subtitle": "子标题"},
    {"image": {"path": "xxx.png", "caption": "图 X 说明"}}
  ],
  "images": [
    {
      "path": "图片文件名",
      "caption": "图片说明文字",
      "note": "备注（可选）"
    }
  ]
}
```

### content 支持的格式

- **纯文本**：直接写字符串
- **子标题**：`{"subtitle": "标题文字"}`（15pt，加粗）
- **分序号**：`{"numbered": "1）标题文字"}`（15pt，加粗，用于粗，用于①②③下的进一步分序号）
- **内嵌图片**：`{"image": {"path": "xxx.png", "caption": "说明"}}`

---

## 最佳实践

### 内容专业性

- 使用领域标准术语（如S11、VSWR、增益、温差等）
- 数据保留适当精度（如温度保留1位小数：0.012K）
- 注明参数来源（如"来源于datasheet典型值"、"COMSOL仿真结果"）

### 图片准备

- 分辨率建议 ≥ 300dpi
- 格式优先：png > jpg > 其他
- 文件名包含序号或关键词便于匹配
- 图片宽度自动设置为15cm，高度按比例

### 直接用于技术报告

周报内容要做到：
- **独立完整**：不依赖周报外的上下文
- **数据准确**：所有数值、图表都有明确来源
- **逻辑自洽**：结论与分析对应，分析与数据对应
- **避免时间词**：不用"本周"、"上周"等，直接陈述事实

### 写作风格

- 客观陈述，避免主观评价
- 使用第三人称或无主语句式
- 结果分析要基于数据，说明趋势和原因
- 总结部分提炼关键发现，不过度推论

---

## 常见问题

### Q: 图片插入失败？

A: 检查图片路径是否正确，确保图片文件存在。支持的格式：png, jpg, jpeg, bmp, tiff, gif

### Q: 中文显示乱码？

A: 确保系统安装了中文字体（宋体、微软雅黑等），配置文件使用UTF-8编码

### Q: 如何修改样式？

A: 可以提供自定义模板文件（`--template`参数），或修改脚本中的 `_setup_styles` 方法

### Q: 图片说明为什么是黑色？

A: 脚本已覆盖Word默认的Caption样式蓝色，设置为黑色以符合中文文档规范

### Q: 模块之间间距太大/太小？

A: 可以修改脚本中的 `MODULE_TITLE_SPACE_BEFORE` 常量（当前为12pt）

---

## 许可证

MIT License
