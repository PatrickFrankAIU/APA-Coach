#!/bin/bash
# Check all help resource URLs in checkApaFormatting.js
# Used by the SessionStart hook. Outputs JSON with systemMessage if any links fail.

REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"
JS_FILE="$REPO_DIR/src/checks/checkApaFormatting.js"

[ -f "$JS_FILE" ] || exit 0

# Extract https:// URLs from url: lines
mapfile -t urls < <(grep -oP "url:\s*['\"]?\K(https://[^'\"]+)" "$JS_FILE" | sort -u)

tmpdir=$(mktemp -d)

for url in "${urls[@]}"; do
  (
    code=$(curl -s --head --connect-timeout 5 --max-time 8 -o /dev/null -w "%{http_code}" -L "$url" 2>/dev/null)
    echo "$code $url" > "$tmpdir/$(echo -n "$url" | md5sum | cut -d' ' -f1).txt"
  ) &
done

wait

failed=()
while IFS= read -r result_file; do
  read -r code url < "$result_file"
  # 200 OK, 301/302/303 redirects, 403 bot protection (Scribbr) = treat as live
  if [[ ! "$code" =~ ^(200|301|302|303|403)$ ]]; then
    failed+=("HTTP $code — $url")
  fi
done < <(find "$tmpdir" -name "*.txt")

rm -rf "$tmpdir"

if [ ${#failed[@]} -gt 0 ]; then
  joined=$(printf '%s\\n' "${failed[@]}")
  printf '{"systemMessage": "⚠️ APA Coach URL check: %d link(s) may be dead:\\n%s"}\n' "${#failed[@]}" "$joined"
fi
