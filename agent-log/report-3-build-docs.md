libghostty-vt build scripts, csproj native copy, and docs are done. Nothing committed. Did not touch ConPtyProcess.cs, Platform/Linux/PtyProcess.cs, InlineProcessManager.cs or Terminal/.

FILES CHANGED
- scripts/build-libghostty.sh (new, chmod +x)
- scripts/build-libghostty.ps1 (new)
- ExcelConsole.csproj
- .gitignore (appended `native/`)
- README.md (new "### Optional: libghostty-vt" subsection under Quick Start; one sentence in "A note on the code")
- CONTRIBUTING.md (new row in ground-rules table)
- CLAUDE.md (Build & Run line; `Terminal/` bullet; InlineProcessManager bullet; Conventions: namespace + optional-native-libs bullet)
- CHANGELOG.md (filled the existing Unreleased section with a "Changed" entry; note: Unreleased sits below 0.36.0, left in place)
- docs/faq.md, docs/tour.md (one sentence each)

CSPROJ: `Compile/None Remove` for native/** and scripts/**; `QsNativeArch` property from RuntimeInformation.OSArchitecture (X64→x64, Arm64→arm64); Windows copies `native/win-$(QsNativeArch)/ghostty-vt.dll`, Linux copies `native/linux-$(QsNativeArch)/libghostty-vt.so*`, flattened to output root. No dylib group (script doesn't support macOS).

BUILD-COPY VERIFICATION
- Dummy native/linux-x64/libghostty-vt.so.1 + decoys in linux-arm64 and win-x64.
- `-getItem:None` listed only linux-x64 file; QsNativeArch=x64.
- `dotnet build`: 0 warnings, 0 errors; bin/Debug/net9.0/libghostty-vt.so.1 copied, decoys not copied.
- Deleted native/, rebuilt: 0/0, no libghostty in bin. Missing native/ doesn't break the build.

`bash -n scripts/build-libghostty.sh`: exit 0.

NOT VERIFIED: neither script run end-to-end (no zig/minisign/PowerShell here). Needs a real run with Zig 0.17 to confirm `--system` flags and the `libghostty-vt.so.1.0.0` output name.
