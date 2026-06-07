# quicksheet-tldr-ext

Show [tldr-pages](https://github.com/tldr-pages/tldr) command cheatsheets directly on your QuickSheet desktop grid. Instant reference for any CLI command without leaving your wallpaper.

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-tldr-ext
```

## Usage

```
tldr: tar
tldr: git-rebase
tldr: ffmpeg
tldr: docker-compose
```

Output fills grid rows with a formatted cheatsheet:

```
Archiving utility.
Often combined with a compression method, such as gzip or bzip2.
▸ Create an archive and write it to a file:
  tar cf {target.tar} {path1} {path2} ...
▸ Create a gzipped archive and write it to a file:
  tar czf {target.tar.gz} {path1} {path2} ...
▸ Extract a (compressed) archive into the current directory:
  tar xf {source.tar[.gz|.bz2|.xz]}
```

## Why

Your desktop wallpaper becomes a living command reference. Pin commonly-forgotten commands (`tldr: rsync`, `tldr: find`, `tldr: awk`) alongside your project launchers and status monitors. No browser tab needed.

Pairs well with:
- `curl:` extension for testing API endpoints
- `gitst:` for repo status
- `docker:` for container health

## Data Source

Fetches directly from [tldr-pages/tldr](https://github.com/tldr-pages/tldr) on GitHub (raw content). No API key required. Results cached for 30 minutes.

## License

MIT
