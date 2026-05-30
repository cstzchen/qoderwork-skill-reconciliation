# QoderWork Skill: 送货单对账单

根据送货单（.doc 文件）自动生成格式统一的对账单 Excel 文件。

## 功能

- 批量读取文件夹中的 .doc 送货单文件
- 自动提取供货方、客户、日期、产品明细等关键数据
- 生成专业格式的对账单 Excel（含表头样式、大写金额、签章区）
- 自动校验数据准确性

## 安装

将 `送货单对账单` 文件夹复制到 QoderWork 技能目录：

```
# Windows
%USERPROFILE%\.qoderwork\skills\送货单对账单\

# macOS / Linux
~/.qoderwork/skills/送货单对账单/
```

确保目录结构如下：

```
送货单对账单/
└── SKILL.md
```

重启 QoderWork 后即可使用。

## 使用方式

在 QoderWork 对话中输入：

> 帮我做对账单，文件夹路径是 F:\xxx\xxx

或直接在技能菜单中选择"送货单对账单"。

## 环境要求

- Windows 系统（需安装 Microsoft Word，用于解析 .doc 文件）
- Python 3.x + openpyxl 库

## License

MIT
