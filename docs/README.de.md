# XivExdUnpacker

Ein auf [Lumina](https://github.com/NotAdam/Lumina) basierendes Tool zum Entpacken von Final Fantasy XIV EXD-Daten als Ersatz für `rawexd` in `SaintCoinach.Cmd`.

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## Schnellstart

1. Archiv unter [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases) herunterladen und entpacken (erfordert [.NET 10](https://dotnet.microsoft.com/download/dotnet/10.0)).
2. `config.yml.example` nach `config.yml` kopieren und Pfade bei Bedarf anpassen.
3. Ausführen:

```bash
# Alle deutschen Tabellen exportieren
XivExdUnpacker --language de
# Für alle Sprachen Action und Item im <hex:...>-Format exportieren und Ausgabe vorher leeren
XivExdUnpacker --language all --sheets Action Item --hexcode --clear
```

## Befehlszeilenargumente

| Argument | Kurz | Beschreibung | Standard |
| --- | --- | --- | --- |
| `--language` | `-l` | Zielsprachen (`en`, `ja`, `de`, `fr`, `chs`, `ko`, `tc`, `all`) | Erforderlich |
| `--sheets` | `-s` | Zu exportierende Tabellen | Alle |
| `--hexcode` | `-x` | Steuer-Tags als `<hex:...>` beibehalten ([Format-Referenz](./STRING_FORMAT.md)) | `false` |
| `--clear` | `-c` | Ausgabeverzeichnis vor dem Export leeren | `false` |
| `--skip-offset` | `-o` | `offset`-Zeile in CSV überspringen | `false` |
| `--help` | `-h` | Hilfe anzeigen | - |

> Nur fertige CSV-Daten benötigt? Siehe [ffxiv-datamining-hexcode-mixed](https://github.com/Souma-Sumire/ffxiv-datamining-hexcode-mixed).

## Entwicklung

```bash
git clone https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
dotnet run -- --language de
```
