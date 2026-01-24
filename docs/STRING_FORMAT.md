# 完整标签对照表

## 基础标签 (0x00-0x0F)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| -------- | ------ | ------------- | ------------ | ------ |
| 0x00 | None | - | - | 无标签 |
| 0x06 | ResetTime | `<hex:020601...03>` | `<ResetTime/>` | 重置时间值 |
| 0x07 | Time | `<hex:020701...03>` | `<Time(...)/>` | 时间设置 |
| 0x08 | If | `<hex:020801...03>` | `<If(condition,true,false)/>` | 条件判断 |
| 0x09 | Switch | `<hex:020901...03>` | `<Switch(value,case1,case2,...)/>` | 多分支选择 |
| 0x0A | Unknown0A | `<hex:020A01...03>` | `<Unknown0A(...)/>` | 未知用途 |
| 0x0C | IfEquals | `<hex:020C01...03>` | `<IfEquals(left,right,true,false)/>` | 相等判断 |

## 格式标签 (0x10-0x1F)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| -------- | ------ | ------------- | ------------ | ------ |
| 0x10 | LineBreak | `<hex:02100103>` | `\n` 或 `<LineBreak/>` | 换行符 |
| 0x12 | Gui | `<hex:021201...03>` | `<Gui(icon)/>` | GUI 图标 (手柄/键盘图标) |
| 0x13 | Color | `<hex:0213F00A03>` | `<Color(10)>` | 颜色开始 |
| 0x13 | Color (关闭) | `<hex:0213EC03>` | `</Color>` | 颜色结束 |
| 0x14 | Unknown14 | `<hex:021401...03>` | `<Unknown14(...)/>` | 未知用途 |
| 0x16 | SoftHyphen | `<hex:021601...03>` | `\u00AD` | 软连字符 |
| 0x17 | Unknown17 | `<hex:021701...03>` | `<Unknown17(...)/>` | 未知 (日文专用) |
| 0x19 | Emphasis2 | `<hex:021901...03>` | `<Emphasis2>` | 强调2 (可能是粗体) |
| 0x1A | Emphasis | `<hex:021A01...03>` | `<Emphasis>` | 强调 (斜体) |
| 0x1D | Indent | `<hex:021D01...03>` | `<Indent(...)/>` | 缩进 |
| 0x1E | CommandIcon | `<hex:021E01...03>` | `<CommandIcon(id)/>` | 命令图标 |
| 0x1F | Dash | `<hex:021F0103>` | `–` | 破折号 |

## 数值标签 (0x20-0x2F)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| -------- | ------ | ------------- | ------------ | ------ |
| 0x20 | Value | `<hex:0220D003>` | `<Value(param1)>` | 参数值 |
| 0x22 | Format | `<hex:022201...03>` | `<Format(...)/>` | 格式化 |
| 0x24 | TwoDigitValue | `<hex:022401...03>` | `<TwoDigitValue(...)/>` | 两位数值 (补零) |
| 0x28 | Sheet | `<hex:0228F2000BF00103>` | `<Sheet(Item,11,1)>` | 表引用 |
| 0x28 | Sheet (关闭) | `<hex:0228EC03>` | `</Sheet>` | 表引用结束 |
| 0x29 | Highlight | `<hex:022901...03>` | `<Highlight>` | 高亮 |
| 0x2B | Clickable | `<hex:022B01...03>` | `<Clickable(...)>` | 可点击 (NPC/物品/玩家) |
| 0x2C | Split | `<hex:022C01...03>` | `<Split(input,sep,index)/>` | 字符串分割 |
| 0x2D | Unknown2D | `<hex:022D01...03>` | `<Unknown2D(...)/>` | 未知用途 |
| 0x2E | Fixed | `<hex:022E01...03>` | `<Fixed(...)/>` | 固定 |
| 0x2F | Unknown2F | `<hex:022F01...03>` | `<Unknown2F(...)/>` | 未知用途 |

## 语言特定表引用 (0x30-0x3F)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| -------- | ------ | ------------- | ------------ | ------ |
| 0x30 | SheetJa | `<hex:023001...03>` | `<SheetJa(...)/>` | 日文表引用 |
| 0x31 | SheetEn | `<hex:023101...03>` | `<SheetEn(...)/>` | 英文表引用 |
| 0x32 | SheetDe | `<hex:023201...03>` | `<SheetDe(...)/>` | 德文表引用 |
| 0x33 | SheetFr | `<hex:023301...03>` | `<SheetFr(...)/>` | 法文表引用 |

## UI 标签 (0x40-0x5F)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| -------- | ------ | ------------- | ------------ | ------ |
| 0x40 | InstanceContent | `<hex:024001...03>` | `<InstanceContent(...)>` | 副本内容 (可点击) |
| 0x48 | UIForeground | `<hex:024801...03>` | `<UIForeground(...)>` | UI 前景色 |
| 0x49 | UIGlow | `<hex:024901...03>` | `<UIGlow(...)>` | UI 发光效果 |
| 0x4A | RubyCharacters | `<hex:024A01...03>` | `<RubyCharacters(...)>` | 注音字符 (主要用于日文) |
| 0x50 | ZeroPaddedValue | `<hex:025001...03>` | `<ZeroPaddedValue(value,width)/>` | 补零数值 |

## 其他标签 (0x60+)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| -------- | ------ | ------------- | ------------ | ------ |
| 0x60 | Unknown60 | `<hex:026001...03>` | `<Unknown60(...)/>` | 未知 (金碟公告前缀) |
