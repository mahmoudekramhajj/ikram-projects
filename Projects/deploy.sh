#!/bin/bash
# deploy.sh — Push code + update stable deployments
# Usage: ./deploy.sh [project_name]  (no args = all projects)

set -e
PROJECTS_DIR="$(cd "$(dirname "$0")" && pwd)"
CONFIG="$PROJECTS_DIR/deploy-config.json"

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

get_deploy_id() {
  python -c "
import json
with open(r'$(cygpath -w "$CONFIG")') as f:
    cfg = json.load(f)
print(cfg['projects'].get('$1', ''))
"
}

deploy_project() {
  local name="$1"
  local dir="$PROJECTS_DIR/$name"

  if [ ! -d "$dir" ]; then
    echo -e "${RED}  ✗ $name — folder not found${NC}"
    return 1
  fi

  local deploy_id=$(get_deploy_id "$name")

  if [ -z "$deploy_id" ]; then
    echo -e "${YELLOW}  ⊘ $name — no deployment ID in config${NC}"
    return 0
  fi

  cd "$dir"
  echo -ne "  $name: push..."
  clasp push -f > /dev/null 2>&1
  echo -ne " ✓ deploy..."
  result=$(clasp deploy -i "$deploy_id" -d "auto $(date +%Y-%m-%d)" 2>&1)
  version=$(echo "$result" | grep -o '@[0-9]*' | head -1)
  echo -e " ${GREEN}✓ $version${NC}"
  cd "$PROJECTS_DIR"
}

WEBAPP_PROJECTS=(
  "Ikram"
  "Airport Search"
  "Mina Camp Search"
  "Hotel Management"
  "Sales Operations Report"
  "Reception Airport"
  "Pilgrim App"
  "Guide App"
)

echo ""
echo "═══════════════════════════════════════"
echo "  Ekram Aldyf — Deploy & Update"
echo "═══════════════════════════════════════"
echo ""

if [ -n "$1" ]; then
  deploy_project "$1"
else
  for proj in "${WEBAPP_PROJECTS[@]}"; do
    deploy_project "$proj"
  done
fi

echo ""
echo -e "${GREEN}Done!${NC} Same URLs — website works instantly."
echo "═══════════════════════════════════════"
