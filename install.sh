#!/usr/bin/env bash
#
# Downloads and installs the Outlook sorter, then hands off to the setup wizard.
# The macOS and Linux twin of install.ps1.
#
#   curl -fsSL https://raw.githubusercontent.com/aaronv02/claude-routine-outlook-sorter/main/install.sh | bash
#
# Or, to choose where it goes:
#
#   curl -fsSL .../install.sh -o install.sh && bash install.sh ~/tools/outlook-sorter
#
# Three things it deliberately does NOT do: install Node (that can need an
# administrator, and silently installing a runtime on someone's machine is not
# this script's business), sign in (the mailbox owner does that herself in the
# wizard), or restart Claude Desktop.

set -euo pipefail

TARGET="${1:-$HOME/outlook-sorter}"
REPO_ZIP='https://github.com/aaronv02/claude-routine-outlook-sorter/archive/refs/heads/main.zip'
INNER_FOLDER='claude-routine-outlook-sorter-main'

bold=$'\033[1m'; dim=$'\033[2m'; green=$'\033[32m'; yellow=$'\033[33m'; reset=$'\033[0m'

step() { printf '\n%s%s%s\n' "$bold" "$1" "$reset"; }
ok()   { printf '  %sOK%s  %s\n' "$green" "$reset" "$1"; }
warn() { printf '  %s!%s   %s\n' "$yellow" "$reset" "$1"; }

printf '\n%sOutlook Sorter - installer%s\n' "$bold" "$reset"

# ---------------------------------------------------------------------------
# Node
# ---------------------------------------------------------------------------

step 'Checking for Node.js...'

if ! command -v node >/dev/null 2>&1; then
  warn 'Node.js is not installed.'
  cat <<'EOF'

  Install the LTS version from https://nodejs.org (or `brew install node` on a Mac
  with Homebrew), then run this installer again.
EOF
  exit 1
fi

NODE_VERSION="$(node --version | sed 's/^v//')"
NODE_MAJOR="${NODE_VERSION%%.*}"
if [ "$NODE_MAJOR" -lt 20 ]; then
  warn "Node $NODE_VERSION is too old - version 20 or newer is required."
  printf '\n  Install the current LTS from https://nodejs.org, then run this again.\n\n'
  exit 1
fi
ok "Node $NODE_VERSION"

command -v npm >/dev/null 2>&1 || { warn 'npm was not found, which is unusual - it ships with Node.'; exit 1; }

# ---------------------------------------------------------------------------
# Where it goes
# ---------------------------------------------------------------------------

step "Installing to $TARGET"

case "$TARGET" in
  *OneDrive*|*"iCloud Drive"*)
    warn 'That path is inside a syncing folder.'
    printf '      Sync will fight with node_modules. Somewhere plainly local is better.\n'
    printf '      Continue anyway? [y/N] '
    read -r answer < /dev/tty
    [[ "$answer" =~ ^[Yy] ]] || exit 1
    ;;
esac

if [ -d "$TARGET" ] && [ -n "$(ls -A "$TARGET" 2>/dev/null)" ]; then
  warn "$TARGET already exists and is not empty."
  printf '      Continuing overwrites the program files. Your sign-in (.env) is kept.\n'
  printf '      Continue? [y/N] '
  read -r answer < /dev/tty
  [[ "$answer" =~ ^[Yy] ]] || exit 1
fi

# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

step 'Downloading...'

TEMP="$(mktemp -d)"
trap 'rm -rf "$TEMP"' EXIT

curl -fsSL "$REPO_ZIP" -o "$TEMP/source.zip"
ok "Downloaded $(( $(wc -c < "$TEMP/source.zip") / 1024 )) KB"

# -q so unzip's file list doesn't scroll the actual instructions off screen.
unzip -q "$TEMP/source.zip" -d "$TEMP"
[ -d "$TEMP/$INNER_FOLDER" ] || { warn "The archive did not contain '$INNER_FOLDER'."; exit 1; }

mkdir -p "$TARGET"
# Copied per-item rather than replacing the directory, so an existing .env and
# routine/.local survive an upgrade. Losing the .env would mean signing in again
# for no reason.
cp -R "$TEMP/$INNER_FOLDER/." "$TARGET/"
ok "Files in place at $TARGET"

# ---------------------------------------------------------------------------
# Dependencies
# ---------------------------------------------------------------------------

step 'Installing dependencies (this takes a minute)...'

cd "$TARGET"
if ! npm install --no-fund --no-audit; then
  warn 'npm install failed.'
  printf '\n  Common causes: no internet, a proxy, or a permissions problem.\n'
  printf '  Retry by hand:\n\n      cd %s\n      npm install\n\n' "$TARGET"
  exit 1
fi
ok 'Dependencies installed'

if [ "${SKIP_SETUP:-}" = "1" ]; then
  printf '\nDone. Run the wizard when ready:\n\n    cd %s\n    npm run setup\n\n' "$TARGET"
  exit 0
fi

# ---------------------------------------------------------------------------
# Hand off
# ---------------------------------------------------------------------------

cat <<EOF

${bold}Starting setup.${reset}

The next part needs the mailbox owner: she signs in partway through, and whoever
signs in is whose mail gets sorted.

EOF

# stdin is redirected explicitly because this script is often piped into bash,
# which leaves the wizard reading the script itself instead of the keyboard - it
# would race through every prompt with garbage and appear to hang.
npm run setup < /dev/tty
