# XivExdUnpacker

A Final Fantasy XIV EXD data unpacking tool powered by [Lumina](https://github.com/NotAdam/Lumina), replacing the `rawexd` workflow of `SaintCoinach.Cmd`.

[English](./README.md) | [日本語](./docs/README.ja.md) | [Deutsch](./docs/README.de.md) | [Français](./docs/README.fr.md) | [简体中文](./docs/README.cn.md) | [한국어](./docs/README.ko.md) | [繁體中文](./docs/README.tc.md)

## Quick Start

1. Download from [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases) ([.NET 10](https://dotnet.microsoft.com/download/dotnet/10.0) required).
2. Copy `config.yml.example` to `config.yml` and adjust paths if needed.
3. Run:

```bash
# Export all English sheets
XivExdUnpacker --language en
# Export all languages' Action and Item sheets in <hex:...> format, clearing output first
XivExdUnpacker --language all --sheets Action Item --hexcode --clear
```

## CLI Options

| Argument | Short | Description | Default |
| --- | --- | --- | --- |
| `--language` | `-l` | Target languages (`en`, `ja`, `de`, `fr`, `chs`, `ko`, `tc`, `all`) | Required |
| `--sheets` | `-s` | Specific sheet names to export | All |
| `--hexcode` | `-x` | Preserve control tags as `<hex:...>` ([format reference](./docs/STRING_FORMAT.md)) | `false` |
| `--clear` | `-c` | Clear output directory before exporting | `false` |
| `--skip-offset` | `-o` | Omit the `offset` row in CSV output | `false` |
| `--help` | `-h` | Show help message | - |

> Looking for unpacked CSVs directly? See [ffxiv-datamining-hexcode-mixed](https://github.com/Souma-Sumire/ffxiv-datamining-hexcode-mixed).

## Development

```bash
git clone https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
dotnet run -- --language en
```
