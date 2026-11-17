# XivExdUnpacker

基于 [Lumina](https://github.com/NotAdam/Lumina) 的《最终幻想XIV》EXD 数据解包工具，用于替代 `SaintCoinach.Cmd` 的 `rawexd` 导出功能。

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## 快速开始

1. 从 [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases) 下载压缩包并解压（需 [.NET 10](https://dotnet.microsoft.com/download/dotnet/10.0)）。
2. 将 `config.yml.example` 复制为 `config.yml`，按需修改路径（设置 `uiLanguage: chs` 可切换中文界面）。
3. 运行：

```bash
# 导出简体中文全部数据表
XivExdUnpacker --language chs
# 导出所有语言的 Action 与 Item 表，保留 <hex:...> 控制码并在导出前清空输出
XivExdUnpacker --language all --sheets Action Item --hexcode --clear
```

## 命令行参数

| 参数 | 简写 | 说明 | 默认值 |
| --- | --- | --- | --- |
| `--language` | `-l` | 目标语言（`en`、`ja`、`de`、`fr`、`chs`、`ko`、`tc`、`all`） | 必填 |
| `--sheets` | `-s` | 指定导出的表名 | 全部 |
| `--hexcode` | `-x` | 保留原始控制码为 `<hex:...>`（参见[标签格式说明](./STRING_FORMAT.md)） | `false` |
| `--clear` | `-c` | 导出前清空输出目录 | `false` |
| `--skip-offset` | `-o` | 跳过 CSV 中的 `offset` 行 | `false` |
| `--help` | `-h` | 显示帮助信息 | - |

> 只需现成的解包 CSV 数据？可直接访问 [ffxiv-datamining-hexcode-mixed](https://github.com/Souma-Sumire/ffxiv-datamining-hexcode-mixed)。

## 本地开发

```bash
git clone https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
dotnet run -- --language chs
```
