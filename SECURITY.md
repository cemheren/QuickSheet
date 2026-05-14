# Security policy

## Reporting a vulnerability

If you find a security issue in QuickSheet, please email **cemheren@gmail.com** rather than opening a public issue. I'll respond within a week.

Useful things to include:
- A short description of the issue.
- Steps to reproduce (a CSV, a key sequence, a specific platform).
- Whether the issue is exploitable from a normal user session or requires special privileges.

## Scope

QuickSheet is a desktop side project that runs as the current user. The threat model is narrow:

- **In scope:** flaws that let an unprivileged process or a hostile CSV escalate to running arbitrary commands as the QuickSheet user (e.g. cell-value parsing bugs that lead to unintended subprocess execution), or that bypass the explicit `r:` / `i:` / `ext:` opt-ins.
- **Also in scope:** flaws in the [extension protocol](docs/extensions.md) (manifest parsing, JSON-line handling, subprocess lifecycle) that could let a malicious extension repo affect QuickSheet beyond what its own subprocess can do.
- **Out of scope:** what a `r:` or `i:` cell does once the user has activated it — those prefixes deliberately run arbitrary commands. Likewise `ext: github:user/repo` deliberately clones and executes a remote repo. Treat extensions like any other software you `git clone && run`.
- **Out of scope:** vulnerabilities in third-party APIs that extensions happen to call.

## Dependencies

QuickSheet's core has **zero NuGet dependencies** as a deliberate policy. Supply-chain vulnerabilities in third-party packages are not a concern for the core. Reference extensions follow the same rule.

## Disclosure

After a fix lands, I'll credit the reporter (with their consent) in the release notes and in the relevant commit message.
