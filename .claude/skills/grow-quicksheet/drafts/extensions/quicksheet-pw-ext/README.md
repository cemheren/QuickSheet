# quicksheet-pw-ext

Secure password generator for [QuickSheet](https://github.com/cemheren/QuickSheet). Generates cryptographically secure passwords, PINs, hex strings, and passphrases directly in a cell. Zero network — all randomness from `System.Security.Cryptography`.

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-pw-ext
```

## Usage

```
pw: 16, 1, 3
```

Parameters: `pw: <spec>, <start_row>, <num_rows>`

### Modes

| Input | Output |
|-------|--------|
| `pw:` | 20-char password (letters+digits+symbols) |
| `pw: 32` | 32-char password |
| `pw: 24 alpha` | 24-char alphanumeric only |
| `pw: 64 hex` | 64-char hex string |
| `pw: pin 8` | 8-digit PIN |
| `pw: 5 words` | 5-word passphrase (diceware-style) |

### Output

```
🔑 x7#Kp!mR2&vQ9@nL4$jZ
20 chars (letters+digits+symbols)
🟢 Very strong (~130 bits)
```

Each activation generates a fresh password — refresh the cell to regenerate.

## Strength indicator

- 🟢 **Very strong** — 128+ bits of entropy
- 🟢 **Strong** — 80–127 bits
- 🟡 **Good** — 60–79 bits
- 🟡 **Fair** — 40–59 bits
- 🔴 **Weak** — below 40 bits

## Security

- Uses `System.Security.Cryptography.RandomNumberGenerator` (CSPRNG)
- Passwords never leave your machine — no network calls
- No logging, no caching, no persistence

## Requirements

- .NET 9 SDK

## License

MIT
