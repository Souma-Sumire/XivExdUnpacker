# XivExdUnpacker

一个基于 [Lumina](https://github.com/NotAdam/Lumina) 的最终幻想 XIV EXD 数据解包工具，用于平替 `SaintCoinach.Cmd` 的 `rawexd` 功能。

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## 使用

### 环境要求

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- FFXIV 的本地安装

### 安装

1. 从 [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases) 下载最新版本。
2. 解压压缩包。

### 配置

将 `config.yml.example` 复制为 `config.yml`，并编辑设置各服的 `path` 和 `outputDir`

### 运行

```bash
# 显示帮助信息
XivExdUnpacker.exe --help

# 导出中文的所有表
XivExdUnpacker.exe --language cn

# 导出英文和日文的 Action 和 Item 表
XivExdUnpacker.exe --language en ja --sheets Action Item

# 导出所有语言，输出 hexcode，导出前清空输出目录
XivExdUnpacker.exe --language all --hexcode --clear
```

### 命令行参数

| 参数 | 简写 | 说明 | 默认值 |
| ---- | ---- | ---- | ------ |
| `--language` | `-l` | 指定要导出的语言 (必需) | - |
| `--sheets` | `-s` | 指定要导出的表名 | 全部 |
| `--hexcode` | `-x` | 保留原始数据 | false |
| `--clear` | `-c` | 导出前清空输出目录 | false |
| `--skip-offset` | - | 跳过 CSV 的 offset 行 | false |
| `--help` | `-h` | 显示帮助信息 | - |

## 开发

### 准备项目

```bash
# 克隆项目
git clone --recursive https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
# 初始化子模块
git submodule update --init --recursive
# 更新 Schema 子模块到最新版本
git submodule update --remote
```

### 运行

```bash
# 显示帮助信息
dotnet run -- --help
# 下略，同上
```
