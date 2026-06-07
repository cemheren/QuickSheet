# quicksheet-jwt-ext

Decode JWT tokens right in your QuickSheet grid. Paste a token, see the header and payload claims rendered inline — with `exp`/`iat`/`nbf` timestamps shown as human-readable dates.

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-jwt-ext
```

## Usage

```
jwt: eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiaWF0IjoxNTE2MjM5MDIyfQ.SflKxwRJSMeKKF2QT4fwpMeJf36POk6yJV_adQssw5c
```

Output fills grid rows below:

```
── Header ──
  alg: HS256
  typ: JWT
── Payload ──
  sub: 1234567890
  name: John Doe
  iat: 1516239022 (2018-01-18 01:30 UTC)
── Sig: present ──
```

## Why

Debugging OAuth/OIDC flows? Checking token expiry? Instead of opening jwt.io or piping through `jq`, decode tokens in-place on your desktop dashboard. Pairs well with the `ping:` and `tls:` extensions for an SRE monitoring panel.

## Protocol

Standard QuickSheet JSON-lines extension protocol. Prefix: `jwt`. See [QuickSheet extension docs](https://github.com/cemheren/QuickSheet#extensions--make-your-desktop-do-more) for details.

## License

MIT
