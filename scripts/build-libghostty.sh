#!/usr/bin/env bash
# Build libghostty-vt from the official, checksum-pinned source tarball and
# install the shared library into native/<rid>/ so `dotnet build` copies it
# next to the QuickSheet executable.
#
# Requires: Zig 0.17.x on PATH, curl, tar, sha256sum (or shasum), and
# optionally minisign (signature check is skipped with a note if absent).
#
# Usage: scripts/build-libghostty.sh
set -euo pipefail

# ---- Pinned release (bump here only) ---------------------------------------
LIBGHOSTTY_VERSION="1.0.0"
SRC_URL="https://release.files.ghostty.org/libghostty-vt/${LIBGHOSTTY_VERSION}/libghostty-vt-${LIBGHOSTTY_VERSION}.tar.gz"
SRC_SHA256="4a6ef758be721923e1969e7b81b443dcc92d8dd3fb5a39600c03a8817a753676"
DEPS_URL="https://release.files.ghostty.org/libghostty-vt/${LIBGHOSTTY_VERSION}/libghostty-vt-${LIBGHOSTTY_VERSION}-deps.tar.gz"
DEPS_SHA256="4d5ea5e448e4c7ea8fb55ccbaa8668f2a97d44f438cee831e2b7ecd162d85c29"
MINISIGN_PUBKEY="RWQlAjJC23149WL2sEpT/l0QKy7hMIFhYdQOFy0Z7z7PbneUgvlsnYcV"
ZIG_REQUIRED_PREFIX="0.17."
ZIG_DOWNLOAD_URL="https://ziglang.org/download/"
# -----------------------------------------------------------------------------

die() { echo "error: $*" >&2; exit 1; }

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

# ---- Platform -> rid --------------------------------------------------------
[[ "$(uname -s)" == "Linux" ]] || die "this script builds the Linux library; on Windows use scripts/build-libghostty.ps1"
case "$(uname -m)" in
    x86_64|amd64)   RID="linux-x64" ;;
    aarch64|arm64)  RID="linux-arm64" ;;
    *) die "unsupported architecture: $(uname -m)" ;;
esac

# ---- Toolchain checks -------------------------------------------------------
for tool in curl tar; do
    command -v "$tool" >/dev/null 2>&1 || die "'$tool' not found on PATH"
done

if command -v zig >/dev/null 2>&1; then
    ZIG_VERSION="$(zig version 2>/dev/null || true)"
else
    ZIG_VERSION=""
fi
if [[ "$ZIG_VERSION" != ${ZIG_REQUIRED_PREFIX}* ]]; then
    die "Zig ${ZIG_REQUIRED_PREFIX}x is required (found: ${ZIG_VERSION:-none}). Install it from ${ZIG_DOWNLOAD_URL}"
fi
echo "Using Zig ${ZIG_VERSION}"

sha256_of() {
    if command -v sha256sum >/dev/null 2>&1; then
        sha256sum "$1" | awk '{print $1}'
    else
        shasum -a 256 "$1" | awk '{print $1}'
    fi
}

verify_sha256() {
    local file="$1" expected="$2"
    local actual
    actual="$(sha256_of "$file")"
    if [[ "$actual" != "$expected" ]]; then
        die "SHA-256 mismatch for $(basename "$file"): expected $expected, got $actual"
    fi
    echo "  sha256 OK: $(basename "$file")"
}

# ---- Work area (cleaned up on exit) -----------------------------------------
WORK="$(mktemp -d "${TMPDIR:-/tmp}/libghostty-build.XXXXXX")"
cleanup() { rm -rf "$WORK"; }
trap cleanup EXIT

STAGE="$WORK/stage"
mkdir -p "$STAGE" "$WORK/zig-cache"

fetch_and_verify() {
    local url="$1" sha="$2" name
    name="$(basename "$url")"
    echo "Downloading ${name}"
    curl -fsSL "$url" -o "$WORK/$name"
    verify_sha256 "$WORK/$name" "$sha"

    if command -v minisign >/dev/null 2>&1; then
        curl -fsSL "${url}.minisig" -o "$WORK/${name}.minisig"
        minisign -V -P "$MINISIGN_PUBKEY" -m "$WORK/$name" >/dev/null
        echo "  minisign OK: ${name}"
    else
        echo "  note: minisign not found; signature skipped (pinned sha256 was checked)"
    fi
}

fetch_and_verify "$SRC_URL" "$SRC_SHA256"
fetch_and_verify "$DEPS_URL" "$DEPS_SHA256"

# Unpack into fresh directories inside the work area.
mkdir "$WORK/src" "$WORK/deps"
tar -xzf "$WORK/libghostty-vt-${LIBGHOSTTY_VERSION}.tar.gz" -C "$WORK/src"
tar -xzf "$WORK/libghostty-vt-${LIBGHOSTTY_VERSION}-deps.tar.gz" -C "$WORK/deps"
SRC_DIR="$WORK/src/libghostty-vt-${LIBGHOSTTY_VERSION}"
DEPS_DIR="$WORK/deps/libghostty-vt-${LIBGHOSTTY_VERSION}-deps"
[[ -d "$SRC_DIR" ]]  || die "unexpected archive layout: missing $SRC_DIR"
[[ -d "$DEPS_DIR" ]] || die "unexpected archive layout: missing $DEPS_DIR"

# ---- Build (fully offline: deps supplied via --system) ----------------------
echo "Building libghostty-vt ${LIBGHOSTTY_VERSION} for ${RID} (this can take a few minutes)"
(
    cd "$SRC_DIR"
    ZIG_GLOBAL_CACHE_DIR="$WORK/zig-cache" zig build \
        --prefix "$STAGE" \
        --system "$DEPS_DIR" \
        -Demit-lib-vt \
        -Doptimize=ReleaseFast \
        -Dcpu=baseline \
        -Dstrip=true
)

# The SONAME symlink (libghostty-vt.so.1) is stable across 1.x; it points at
# the fully versioned file, so this works for any patch release.
LIB_REAL="$STAGE/lib/libghostty-vt.so.1"
[[ -e "$LIB_REAL" ]] || die "build did not produce $LIB_REAL"

# ---- Install ----------------------------------------------------------------
DEST_DIR="$REPO_ROOT/native/$RID"
mkdir -p "$DEST_DIR"
# cp -L dereferences symlinks so the tree holds one real file, no links.
cp -L "$LIB_REAL" "$DEST_DIR/libghostty-vt.so.1"
chmod 0644 "$DEST_DIR/libghostty-vt.so.1"

echo
echo "Installed: $DEST_DIR/libghostty-vt.so.1"
echo "Next: run 'dotnet build ExcelConsole.csproj' so the library is copied to the output directory."
echo "QuickSheet will then use it automatically (or set QUICKSHEET_LIBGHOSTTY to a full path)."
