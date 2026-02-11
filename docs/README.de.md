# XivExdUnpacker

Ein auf [Lumina](https://github.com/NotAdam/Lumina) basierendes Tool zum Entpacken von Final Fantasy XIV EXD-Daten, konzipiert als Ersatz für die `rawexd`-Funktionalität von `SaintCoinach.Cmd`.

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## Verwendung

### Anforderungen

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- Lokale Installation von FFXIV

### Installation

1. Laden Sie die neueste Version von [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases) herunter.
2. Entpacken Sie das Archiv.

### Konfiguration

Kopieren Sie `config.yml.example` nach `config.yml` und bearbeiten Sie `path` und `outputDir` für jeden Server.

### Ausführen

```bash
# Hilfe anzeigen
XivExdUnpacker.exe --help

# Alle Tabellen für Chinesisch exportieren
XivExdUnpacker.exe --language cn

# Tabellen Action und Item für Englisch und Japanisch exportieren
XivExdUnpacker.exe --language en ja --sheets Action Item

# Alle Sprachen exportieren, Hexcode ausgeben, Ausgabeverzeichnis vor dem Export leeren
XivExdUnpacker.exe --language all --hexcode --clear
```

### Befehlszeilenargumente

| Argument | Kurz | Beschreibung | Standard |
| ---- | ---- | ---- | ------ |
| `--language` | `-l` | Zu exportierende Sprachen angeben (Erforderlich) | - |
| `--sheets` | `-s` | Namen der zu exportierenden Tabellen angeben | Alle |
| `--hexcode` | `-x` | Rohdaten beibehalten | false |
| `--clear` | `-c` | Ausgabeverzeichnis vor dem Export leeren | false |
| `--skip-offset` | `-o` | CSV-Offset-Zeilen überspringen | false |
| `--help` | `-h` | Hilfeinformationen anzeigen | - |

> ### Sie wollen nur die Daten?
>
> Schauen Sie sich [ffxiv-datamining-hexcode-mixed](https://github.com/Souma-Sumire/ffxiv-datamining-hexcode-mixed) an.

## Entwicklung

```bash
git clone https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
dotnet run -- --language de
```
