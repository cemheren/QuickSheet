# quicksheet-subnet-ext

CIDR/subnet calculator for [QuickSheet](https://github.com/cemheren/QuickSheet). Displays network ranges, broadcast address, host count, and netmask — right on your desktop wallpaper.

## Install

Add to any cell in QuickSheet:

```
ext: github:cemheren/quicksheet-subnet-ext
```

## Usage

```
subnet: 192.168.1.0/24    → Network, broadcast, 254 hosts, range
subnet: 10.0.0.0/16       → 65,534 usable hosts
subnet: 172.16.0.0/12     → Class B private range breakdown
subnet: 203.0.113.42/28   → Small allocation (14 hosts)
subnet: 10.1.1.1          → Single host (/32)
```

## Output

| Row | Content |
|-----|---------|
| 1   | 🌐 `<normalized CIDR>` |
| 2   | Network address |
| 3   | Broadcast address |
| 4   | Usable host range (first – last) |
| 5   | Host count |
| 6   | Subnet mask |

## Why

SREs and netops engineers reference subnet boundaries constantly — firewall rules, VPC planning, incident response. Instead of alt-tabbing to an online calculator, glance at your wallpaper. Pairs well with `ping:`, `tls:`, `dns:`, and `ip:` extensions for a complete network ops dashboard.

## Features

- Standard CIDR notation (`10.0.0.0/24`)
- Bare IP defaults to `/32`
- RFC 3021 point-to-point (`/31`) handled correctly
- Human-readable host counts with thousands separators
- Zero external dependencies

## Requirements

- .NET 9+ runtime
- QuickSheet with extension protocol v1+

## License

MIT
