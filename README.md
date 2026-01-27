# XivExdUnpacker

A Final Fantasy XIV EXD data unpacking tool based on [Lumina](https://github.com/NotAdam/Lumina), intended as a replacement for the `rawexd` functionality of `SaintCoinach.Cmd`.

[English](./README.md) | [日本語](./docs/README.ja.md) | [Deutsch](./docs/README.de.md) | [Français](./docs/README.fr.md) | [简体中文](./docs/README.cn.md) | [한국어](./docs/README.ko.md) | [繁體中文](./docs/README.tc.md)

## Usage

### Requirements

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- Local installation of FFXIV

### Installation

1. Download the latest release from [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases).
2. Extract the archive.

### Configuration

Copy `config.yml.example` to `config.yml` and edit the `path` and `outputDir` for each server.

### Run

```bash
# Show help
XivExdUnpacker.exe --help

# Export all tables for Chinese
XivExdUnpacker.exe --language cn

# Export Action and Item tables for English and Japanese
XivExdUnpacker.exe --language en ja --sheets Action Item

# Export all languages, output hexcode, clear output directory before exporting
XivExdUnpacker.exe --language all --hexcode --clear
```

### Command Line Arguments

| Argument | Short | Description | Default |
| ---- | ---- | ---- | ------ |
| `--language` | `-l` | Specify languages to export (Required) | - |
| `--sheets` | `-s` | Specify sheet names to export | All |
| `--hexcode` | `-x` | Keep raw data | false |
| `--clear` | `-c` | Clear output directory before exporting | false |
| `--skip-offset` | - | Skip CSV offset rows | false |
| `--help` | `-h` | Show help information | - |

## Development

### Setup

```bash
# Clone the repository
git clone --recursive https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
# Initialize submodules
git submodule update --init --recursive
# Update Schema submodule to latest
git submodule update --remote
```

### Build & Run

```bash
# Show help
dotnet run -- --help
# Same as above
```
