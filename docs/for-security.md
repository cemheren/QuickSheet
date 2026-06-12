# QuickSheet for Security Engineers

If you spend your day rotating certificates, scanning subnets, verifying DNS records,
decoding JWTs, or generating hashes for integrity checks — **your desktop can be a
passive recon board.**

> Already use Burp Suite, Nmap, Wireshark, or CyberChef? QuickSheet is the persistent
> reference layer behind your windows — hostnames, cert expiry dates, hash lookups, and
> quick-decode tools that never leave your screen.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/security-dashboard.csv
```

The starter sheet at `examples/security-dashboard.csv` gives you a deployable frame. Edit
cells live; grid autosaves to CSV every 5 seconds. Add it to your session startup and the
reference board is there every boot.

## What goes on the wallpaper

| Row/column              | What it shows                                             | Extension                                                              |
|-------------------------|-----------------------------------------------------------|------------------------------------------------------------------------|
| **TLS expiry**          | Days remaining on each certificate — red when < 30        | [`tls`](https://github.com/Deskworks/quicksheet-tls-ext)               |
| **DNS records**         | A / AAAA / MX / NS records for any domain                | [`dns`](https://github.com/cemheren/quicksheet-dns-ext)                |
| **Subnet calculator**   | Network, broadcast, host range for a CIDR                 | [`subnet`](https://github.com/cemheren/quicksheet-subnet-ext)          |
| **Port probe**          | TCP open/closed for well-known ports                      | [`portck`](https://github.com/Deskworks/quicksheet-portck)             |
| **Latency / reachability** | Round-trip time to hosts and DNS resolvers             | [`ping`](https://github.com/cemheren/quicksheet-ping-ext)              |
| **Traceroute**          | Hop-by-hop path to a target                               | [`tracert`](https://github.com/cemheren/quicksheet-tracert)            |
| **IP info**             | Geo, ASN, ISP for any IP address                          | [`ip`](https://github.com/cemheren/quicksheet-ip-ext)                  |
| **JWT decode**          | Header + payload of a JWT without leaving your desktop    | [`jwt`](https://github.com/cemheren/quicksheet-jwt-ext)                |
| **Hash generation**     | SHA-256, MD5, SHA-1 of any input string                   | [`hashgen`](https://github.com/cemheren/quicksheet-hashgen)            |
| **Password generation** | Cryptographically random passwords on demand              | [`pw`](https://github.com/cemheren/quicksheet-pw-ext)                  |

All extensions install with a single cell: `ext: github:<owner>/<repo>`. No package
manager. No daemon config. QuickSheet clones the repo, starts the subprocess, and the
prefix is live.

## Daily security workflows

### Certificate rotation tracking

Keep a column of `tls:` cells pointed at every domain you manage. Glance at your desktop
to spot which certs are approaching expiry. Pair with a `L:` loop cell to re-check every
30 minutes automatically.

```
tls: example.com
tls: api.example.com
tls: mail.example.com
tls: vpn.example.com
```

### Quick recon row

One row of DNS + subnet + port probes gives you instant context when triaging an alert:

```
dns: suspicious-domain.xyz
subnet: 192.168.1.0/24
portck: 192.168.1.1:22
ping: 192.168.1.1
```

### Token inspection

Paste a JWT into a cell prefixed with `jwt:` to decode it on the spot — no browser tabs,
no CyberChef detours. Useful during incident response when you need to verify token claims
quickly.

### Hash verification

Verifying file integrity? Use `hashgen:` with the expected input to generate a hash and
visually compare. Keep a reference column of known-good hashes next to computed values.

### Password generation for service accounts

Need a quick password for a new service account? `pw: 32` generates a 32-character random
password right on your desktop. Copy it, use it, and the cell stays there as a reminder of
what was generated.

## Pair with inline commands

Beyond extensions, prefix any cell with `i:` for live subprocess output:

```
i: dig +short example.com
i: whois -h whois.iana.org example.com | head -5
i: nmap -F 192.168.1.1 2>/dev/null | grep open
i: curl -sI https://example.com | grep -i "strict-transport"
i: openssl s_client -connect example.com:443 </dev/null 2>/dev/null | openssl x509 -noout -dates
```

These run as subprocesses and stream output back into the cell — a live terminal embedded
in your desktop grid.

## Links

- [Main README](../README.md)
- [60-second tour](tour.md)
- [Keyboard shortcuts](keyboard-shortcuts.md)
- [Extension directory](extensions.md)
- [SRE & DevOps guide](for-sre.md) (complementary — ops-focused rather than security-focused)
