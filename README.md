# XivExdUnpacker

一个基于 [Lumina](https://github.com/NotAdam/Lumina) 的最终幻想 XIV EXD 数据解包工具，用于平替 `SaintCoinach.Cmd` 的 `rawexd` 功能。

## 环境要求

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- FFXIV 的本地安装

## 快速开始

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

### 配置

将 `config.yml.example` 复制为 `config.yml`，并编辑设置各服的 `path` 和 `outputDir`

### 运行

```bash
dotnet run
```
