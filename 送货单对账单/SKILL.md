---
name: 送货单对账单
version: 1.0.0
description: Generate reconciliation statements from delivery notes (.doc files). Use when the user wants to create a reconciliation statement, summarize delivery notes, or generate accounts receivable documents.
description_zh: 根据送货单（.doc文件）批量生成对账单Excel文件。适用于用户需要制作对账单、汇总送货记录、生成应收款文件等场景。
user-invocable: true
argument-hint: 指定包含送货单文件的文件夹路径
---

# 送货单对账单制作

根据指定文件夹中的送货单（.doc 格式）文件，自动提取数据并生成格式统一的对账单 Excel 文件。

## 触发条件

用户提到以下关键词时自动触发：对账单、对账、送货单汇总、应收对账、客户对账。

## 工作流程

### 第一步：确认输入

1. 确认用户提供的文件夹路径（包含 `.doc` 送货单文件）
2. 列出文件夹中的所有 `.doc` 文件，向用户确认
3. 对账期间默认取送货单中最早和最晚日期，精确到日（格式：YYYY年M月D日 — YYYY年M月D日）

### 第二步：读取 .doc 文件

.doc 文件是 OLE2 二进制格式，需要借助 Microsoft Word COM 对象转换为文本。

**Windows 环境下的转换方法（PowerShell 脚本）：**

```powershell
$files = @('file1.doc', 'file2.doc')  # 遍历文件夹中所有 .doc
$outDir = '<工作目录>'

foreach ($f in $files) {
    $inPath = "<源文件夹>\$f"
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($f)
    $outPath = "$outDir\$baseName.txt"
    
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($inPath, $false, $true)
        if ($doc) {
            $doc.SaveAs($outPath, 2)  # 2 = wdFormatText
            $doc.Close($false)
        }
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    } catch {
        Write-Output "ERROR: $f - $($_.Exception.Message)"
        try { $word.Quit() } catch {}
    }
    Start-Sleep -Seconds 2
}
```

**重要：每个文件使用独立的 Word 实例，避免 COM 崩溃影响后续文件。**

转换完成后，使用 GBK 编码读取 .txt 文件：

```python
with open(path, 'r', encoding='gbk') as f:
    content = f.read()
```

### 第三步：解析送货单数据

从文本中提取以下字段：

| 字段 | 说明 | 示例 |
|------|------|------|
| 供货方 | 送货单抬头公司名称 | 苏州中精自动化科技有限公司 |
| 客户 | "客户："后面的公司名称 | 宁波中锐电力科技有限公司 |
| 日期 | "日期："后面的日期 | 2026年3月18日 |
| 产品名称 | 表格中的名称列 | 底壳测试组件 |
| 单位 | 套/个/件等 | 套 |
| 数量 | 数字 | 2 |
| 单价 | 数字 | 1500 |
| 金额 | 数量 × 单价 | 3000 |
| 备注 | 快递单号等信息 | 顺丰SF1566275642679 |

**解析注意事项：**
- 一张送货单可能包含多行产品明细
- "以下空" 表示明细结束
- 收货方/送货方签章行不包含产品数据
- 自动识别供货方和客户信息（通用版，不锁定特定公司）

### 第四步：生成对账单 Excel

使用 Python `openpyxl` 库生成 Excel 文件。如未安装，执行 `python -m pip install openpyxl`。

#### 表格结构

```
行1: 标题 "对  账  单"（合并A-H列，居中）
行2: 供货方信息（A-C列） | 客户信息（D-H列）
行3: 电话（A-C列）       | 对账期间（D-H列，精确到日）
行4: 空行（高度8）
行5: 表头（序号/送货日期/产品名称/单位/数量/单价（元）/金额（元）/备注）
行6+: 数据行（按日期排序，同一送货单的多个产品逐行排列）
合计行: 合并A-D列写"合  计"，E列总数量，G列总金额
大写行: 合并A-H列写"合计金额（大写）：XX元整"
空行
签章行: 供货方签章（A-D列） | 客户确认签章（E-H列）
日期行: 日期：    年    月    日（A-D列） | 日期：    年    月    日（E-H列）
```

#### 样式规范

| 元素 | 字体 | 大小 | 加粗 | 对齐 | 填充色 |
|------|------|------|------|------|--------|
| 标题 | 微软雅黑 | 16 | 是 | 居中 | 无 |
| 信息行 | 微软雅黑 | 11 | 否 | 左对齐 | 无 |
| 表头 | 微软雅黑 | 11 | 是 | 居中 | #4472C4（蓝色），字色白色 |
| 数据行 | 微软雅黑 | 11 | 否 | 见下方 | 无 |
| 合计行 | 微软雅黑 | 12 | 是 | 居中/右对齐 | #D9E2F3（浅蓝） |

**数据行对齐方式：**
- 序号、送货日期、单位、数量：居中
- 产品名称、备注：左对齐
- 单价、金额：右对齐，数字格式 `#,##0`

**列宽设定：**
- A（序号）: 6
- B（送货日期）: 14
- C（产品名称）: 28
- D（单位）: 8
- E（数量）: 8
- F（单价）: 12
- G（金额）: 14
- H（备注）: 22

**行高设定：**
- 标题行: 40
- 信息行: 25
- 空行: 8
- 表头行: 30
- 数据行: 26
- 合计行: 32
- 大写行: 30
- 签章行: 50
- 日期行: 30

**边框：** 表头行、数据行、合计行所有单元格均使用细线边框（`Side(style='thin')`）。

#### 大写金额转换

使用以下规则将数字金额转为中文大写：

```
数字: 零壹贰叁肆伍陆柒捌玖
整数单位: 拾佰仟万亿
小数单位: 角分
示例: 22300 → 贰万贰仟叁佰元整
```

#### 输出文件

- **文件名格式：** `对账单_{客户简称}_{起始日期}-{结束日期}.xlsx`
  - 客户简称：取客户公司名称中的核心词（如"宁波中锐电力科技有限公司" → "宁波中锐"）
  - 日期格式：YYYY年M月D日（如"2026年3月18日-5月6日"）
- **保存位置：** 与送货单源文件同一目录

### 第五步：数据校验

生成后必须执行校验：
1. 逐张送货单核对：计算金额 = 数量 × 单价，与原始金额比对
2. 汇总核对：所有明细金额之和 = 合计行金额
3. 输出校验结果，确认无误后交付

### 第六步：交付

使用 `present_files` 工具将文件呈现给用户，并提供简要汇总：
- 送货单份数
- 产品明细项数
- 总数量和总金额
- 对账期间
