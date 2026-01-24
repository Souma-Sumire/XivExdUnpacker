# XivExdUnpacker

Un outil de désarchivage de données EXD de Final Fantasy XIV basé sur [Lumina](https://github.com/NotAdam/Lumina), conçu pour remplacer la fonctionnalité `rawexd` de `SaintCoinach.Cmd`.

[English](../README.md) | [日本語](./README.ja.md) | [Deutsch](./README.de.md) | [Français](./README.fr.md) | [简体中文](./README.cn.md) | [한국어](./README.ko.md) | [繁體中文](./README.tc.md)

## Prérequis

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- Installation locale de FFXIV

## Démarrage rapide

### Préparer le projet

```bash
# Cloner le dépôt
git clone --recursive https://github.com/Souma-Sumire/XivExdUnpacker.git
cd XivExdUnpacker
# Initialiser les sous-modules
git submodule update --init --recursive
# Mettre à jour le sous-module Schema vers la dernière version
git submodule update --remote
```

### Configuration

Copiez `config.yml.example` en `config.yml` et modifiez le `path` et `outputDir` pour chaque serveur.

### Exécution

```bash
# Afficher l'aide
dotnet run -- --help

# Exporter toutes les tables pour le chinois (décodage de chaîne par défaut)
dotnet run -- --language cn

# Exporter les tables Action et Item pour l'anglais et le japonais
dotnet run -- --language en ja --sheets Action Item

# Exporter toutes les langues, conserver les données brutes, vider le répertoire de sortie
dotnet run -- --language all --hexcode --clear

# Utiliser les commandes courtes : exporter le chinois, vider la sortie, utiliser le HEX brut
dotnet run -- -l cn -c -x
```

### Arguments de la ligne de commande

| Argument | Court | Description | Par défaut |
| ---- | ---- | ---- | ------ |
| `--language` | `-l` | Spécifier les langues à exporter (Requis) | - |
| `--sheets` | `-s` | Spécifier les noms des feuilles à exporter | Tout |
| `--hexcode` | `-x` | Conserver les données brutes | false |
| `--clear` | `-c` | Vider le répertoire de sortie avant l'exportation | false |
| `--skip-offset` | - | Sauter les lignes d'offset CSV | false |
| `--help` | `-h` | Afficher l'aide | - |
