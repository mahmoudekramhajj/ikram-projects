#!/bin/bash
# deploy.sh — Push code + update stable deployments + verify + track
# Usage: ./deploy.sh [project_name]  (no args = all webapp projects)

set -e
PROJECTS_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(dirname "$PROJECTS_DIR")"
MANIFEST="$REPO_ROOT/deployments.json"
INDEX_HTML="$REPO_ROOT/index.html"

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
CYAN='\033[0;36m'
NC='\033[0m'

get_deploy_id() {
  python -c "
import json
with open(r'$(cygpath -w "$MANIFEST")') as f:
    cfg = json.load(f)
proj = cfg.get('projects', {}).get('$1', {})
print(proj.get('deploymentId') or '')
"
}

update_manifest() {
  local name="$1"
  local version="$2"
  local timestamp="$3"
  python -c "
import json
path = r'$(cygpath -w "$MANIFEST")'
with open(path) as f:
    cfg = json.load(f)
proj = cfg.get('projects', {}).get('$name', {})
if proj:
    proj['lastVersion'] = '$version'
    proj['lastDeployed'] = '$timestamp'
    with open(path, 'w') as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)
    print('updated')
else:
    print('not_found')
"
}

verify_url() {
  local deploy_id="$1"
  local url="https://script.google.com/macros/s/${deploy_id}/exec"
  local status=$(curl -s -o /dev/null -w "%{http_code}" -L --max-time 15 "$url" 2>/dev/null)
  echo "$status"
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
    echo -e "${YELLOW}  ⊘ $name — no deployment ID in manifest${NC}"
    return 0
  fi

  cd "$dir"

  # Step 1: Push
  echo -ne "  $name: push..."
  clasp push -f > /dev/null 2>&1
  echo -ne " ${GREEN}✓${NC}"

  # Step 2: Deploy
  echo -ne " deploy..."
  local result=$(clasp deploy -i "$deploy_id" -d "auto $(date +%Y-%m-%d)" 2>&1)
  local version=$(echo "$result" | grep -o '@[0-9]*' | head -1)
  echo -ne " ${GREEN}✓ $version${NC}"

  # Step 3: Verify URL
  echo -ne " verify..."
  local status=$(verify_url "$deploy_id")
  if [ "$status" = "200" ] || [ "$status" = "302" ]; then
    echo -ne " ${GREEN}✓ $status${NC}"
  else
    echo -ne " ${RED}✗ $status${NC}"
  fi

  # Step 4: Update manifest
  local timestamp=$(date -u +%Y-%m-%dT%H:%M:%SZ)
  local ver_num=$(echo "$version" | tr -d '@')
  update_manifest "$name" "$ver_num" "$timestamp" > /dev/null

  echo ""
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
  "Transport Management"
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
