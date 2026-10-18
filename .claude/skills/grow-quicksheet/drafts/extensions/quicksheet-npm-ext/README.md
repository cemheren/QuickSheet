# quicksheet-npm-ext

npm registry package lookup extension for [QuickSheet](https://github.com/cemheren/QuickSheet).

Monitor your project's key dependencies at a glance — see latest versions, descriptions, and weekly download counts right on your desktop wallpaper.

## Install

Add to any cell in QuickSheet:

```
ext: github:cemheren/quicksheet-npm-ext
```

## Usage

```
npm: express
npm: react
npm: typescript
npm: @types/node
```

### Output (3-row cell)

```
📦 express
v5.1.0
Fast, unopinionated, minimalist web framework
```

### Output (5-row cell)

```
📦 express
v5.1.0
Fast, unopinionated, minimalist web framework
⬇ 35.2M/wk
MIT
```

## How It Works

Queries the public [npm registry API](https://registry.npmjs.org/) and the
[npm downloads API](https://api.npmjs.org/) — no authentication needed, no
API key, zero NuGet dependencies.

## Requirements

- .NET 9 SDK
- Internet connection (for registry queries)

## License

MIT
