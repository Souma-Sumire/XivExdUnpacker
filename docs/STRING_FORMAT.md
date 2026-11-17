# 完整标签对照表

## 基础标签 (0x00-0x0F)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| ------ | ---- | ------------ | ------------ | ---- |
| 0x06 | SetResetTime | `<hex:020601...03>` | `<setresettime(...)>` | 重置时间值 |
| 0x07 | SetTime | `<hex:020701...03>` | `<settime(...)>` | 设置时间值 |
| 0x08 | If | `<hex:020801...03>` | `<if(cond,true,false)>` | 条件判断 |
| 0x09 | Switch | `<hex:020901...03>` | `<switch(val,case1,case2,...)>` | 多分支选择 |
| 0x0A | PcName | `<hex:020A01...03>` | `<pcname(id)>` | 玩家角色名 |
| 0x0B | IfPcGender | `<hex:020B01...03>` | `<ifpcgender(id,m,f)>` | 角色性别条件判断 |
| 0x0C | IfPcName | `<hex:020C01...03>` | `<ifpcname(id,name,true,false)>` | 角色名匹配判断 |
| 0x0D | Josa | `<hex:020D01...03>` | `<josa(...)>` | 韩语助词判定 |
| 0x0E | Josaro | `<hex:020E01...03>` | `<josaro(...)>` | 韩语助词判定（ro） |
| 0x0F | IfSelf | `<hex:020F01...03>` | `<ifself(id,true,false)>` | 是否为玩家自身判断 |

## 格式标签 (0x10-0x1F)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| ------ | ---- | ------------ | ------------ | ---- |
| 0x10 | NewLine | `<hex:02100103>` | `<br/>` | 换行符 |
| 0x11 | Wait | `<hex:021101...03>` | `<wait(sec)>` | 等待时长 |
| 0x12 | Icon | `<hex:021201...03>` | `<icon(id)>` | GUI 图标（手柄/键盘图标） |
| 0x13 | Color | `<hex:0213F00A03>` | `<color(...)>` | 文本前景色 |
| 0x14 | EdgeColor | `<hex:021401...03>` | `<edgecolor(...)>` | 文本描边色 |
| 0x15 | ShadowColor | `<hex:021501...03>` | `<shadowcolor(...)>` | 文本阴影色 |
| 0x16 | SoftHyphen | `<hex:021601...03>` | `<->` | 软连字符 |
| 0x17 | Key | `<hex:021701...03>` | `<key>` | 按键标识（日文等） |
| 0x18 | Scale | `<hex:021801...03>` | `<scale(n)>` | 字体缩放 |
| 0x19 | Bold | `<hex:021901...03>` | `<bold(bool)>` | 粗体开关 |
| 0x1A | Italic | `<hex:021A01...03>` | `<italic(bool)>` | 斜体开关 |
| 0x1B | Edge | `<hex:021B01...03>` | `<edge(n,n)>` | 描边设置 |
| 0x1C | Shadow | `<hex:021C01...03>` | `<shadow(n,n)>` | 阴影设置 |
| 0x1D | NonBreakingSpace | `<hex:021D01...03>` | `<nbsp/>` | 不换行空格 |
| 0x1E | Icon2 | `<hex:021E01...03>` | `<icon2(id)>` | 扩展图标 |
| 0x1F | Hyphen | `<hex:021F0103>` | `<-->` | 连字符 |

## 数值与文本处理标签 (0x20-0x2F)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| ------ | ---- | ------------ | ------------ | ---- |
| 0x20 | IntegerParameter | `<hex:0220D003>` | `<num(expr)>` | 整数参数值 |
| 0x22 | Format | `<hex:022201...03>` | `<format(expr,hex)>` | 格式化 |
| 0x24 | TwoDigitValue | `<hex:022401...03>` | `<twodigit(expr)>` | 两位数值（补零） |
| 0x27 | Link | `<hex:022701...03>` | `<link(...)>` | 可点击链接 |
| 0x28 | Sheet | `<hex:02280EFF09>ClassJob<hex:E8020103>` | `<sheet(name,row,col)>` | 数据表引用 |
| 0x29 | String | `<hex:022901...03>` | `<string(expr)>` | 字符串表达式 |
| 0x2A | Caps | `<hex:022A01...03>` | `<caps(expr)>` | 全部大写 |
| 0x2B | Head | `<hex:022B01...03>` | `<head(expr)>` | 首字母大写 |
| 0x2C | Split | `<hex:022C01...03>` | `<split(str,sep,idx)>` | 字符串分割 |
| 0x2D | HeadAll | `<hex:022D14FF1202280EFF09>ClassJob<hex:E8021F0303>` | `<headall(expr)>` | 每个单词首字母大写 |
| 0x2E | Fixed | `<hex:022E01...03>` | `<fixed(...)>` | 固定文本 |
| 0x2F | Lower | `<hex:022F01...03>` | `<lower(expr)>` | 全部小写 |

## 语言特定表引用 (0x30-0x3F)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| ------ | ---- | ------------ | ------------ | ---- |
| 0x30 | JaNoun | `<hex:023001...03>` | `<janoun(...)>` | 日文名词表引用 |
| 0x31 | EnNoun | `<hex:023101...03>` | `<ennoun(...)>` | 英文名词表引用 |
| 0x32 | DeNoun | `<hex:023201...03>` | `<denoun(...)>` | 德文名词表引用 |
| 0x33 | FrNoun | `<hex:023301...03>` | `<frnoun(...)>` | 法文名词表引用 |
| 0x34 | ChNoun | `<hex:023401...03>` | `<chnoun(...)>` | 中文名词表引用 |

## UI 标签 (0x40-0x5F)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| ------ | ---- | ------------ | ------------ | ---- |
| 0x40 | LowerHead | `<hex:024001...03>` | `<lowerhead(expr)>` | 首字母小写 |
| 0x41 | SheetSub | `<hex:024101...03>` | `<sheetsub(...)>` | 子表引用 |
| 0x42 | SwitchPlatform | `<hex:024201...03>` | `<switchplatform(...)>` | 按平台切换文本 |
| 0x48 | ColorType | `<hex:024801...03>` | `<colortype(id)>` | UIColor 前景色 |
| 0x49 | EdgeColorType | `<hex:024901...03>` | `<edgecolortype(id)>` | UIColor 描边色 |
| 0x4A | Ruby | `<hex:024A01...03>` | `<ruby(text,ruby)>` | 注音字符（主要用于日文） |
| 0x50 | Digit | `<hex:025001...03>` | `<digit(val,width)>` | 补零数值 |
| 0x51 | Ordinal | `<hex:025101...03>` | `<ordinal(val)>` | 英文序数词 |

## 其他标签 (0x60+)

| Tag ID | 名称 | hexcode 示例 | 人类可读格式 | 说明 |
| ------ | ---- | ------------ | ------------ | ---- |
| 0x60 | Sound | `<hex:026001...03>` | `<sound(jingle,id)>` | 音效指令 |
| 0x61 | LevelPos | `<hex:026101...03>` | `<levelpos(id)>` | 地图位置与坐标 |
