# XivExdUnpacker

Un outil de désarchivage de données EXD de Final Fantasy XIV basé sur [Lumina](https://github.com/NotAdam/Lumina), conçu pour remplacer la fonctionnalité `rawexd` de `SaintCoinach.Cmd`.

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## Utilisation

### Prérequis

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- Installation locale de FFXIV

### Installation

1. Téléchargez la dernière version depuis [Releases](https://github.com/Souma-Sumire/XivExdUnpacker/releases).
2. Extrayez l'archive.

### Configuration

Copiez `config.yml.example` en `config.yml` et modifiez le `path` et `outputDir` pour chaque serveur.

### Exécution

```bash
# Afficher l'aide
XivExdUnpacker.exe --help

# Exporter toutes les tables pour le chinois
XivExdUnpacker.exe --language cn

# Exporter les tables Action et Item pour l'anglais et le japonais
XivExdUnpacker.exe --language en ja --sheets Action Item

# Exporter toutes les langues, sortir le hexcode, vider le répertoire de sortie avant l'exportation
XivExdUnpacker.exe --language all --hexcode --clear
```

### Arguments de la ligne de commande

| Argument | Court | Description | Par défaut |
| ---- | ---- | ---- | ------ |
| `--language` | `-l` | Spécifier les langues à exporter (Requis) | - |
| `--sheets` | `-s` | Spécifier les noms des feuilles à exporter | Tout |
| `--hexcode` | `-x` | Conserver les données brutes | false |
| `--clear` | `-c` | Vider le répertoire de sortie avant l'exportation | false |
| `--skip-offset` | `-o` | Sauter les lignes d'offset CSV | false |
| `--help` | `-h` | Afficher l'aide | - |

> ### Vous voulez juste les données ?
>
> Consultez [ffxiv-datamining-hexcode-mixed](https://github.com/Souma-Sumire/ffxiv-datamining-hexcode-mixed).

## Développement

```bash
git clone https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
dotnet run -- --language fr
```
