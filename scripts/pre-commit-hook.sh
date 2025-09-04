#!/bin/bash
# Pre-commit hook to update headers in root-level .js files with filename, last update date/time, and SHA256 hash

# Function to update or add header in a file
update_header() {
  local file="$1"
  local date="$2"
  local filename=$(basename "$file")
  # Compute SHA256 from line 2 onward (excluding header)
  local hash=$(tail -n +2 "$file" | shasum -a 256 | awk '{print $1}')
  # Check if first line matches the header pattern
  if head -1 "$file" | grep -q "^/\* $filename Last Update"; then
    # Replace the first line
    sed -i '' "1s|.*|/* $filename Last Update $date <$hash>|" "$file"
  else
    # Insert the header at the top
    { echo "/* $filename Last Update $date <$hash>"; cat "$file"; } > temp && mv temp "$file"
  fi
}

# Process each .js file in the root directory
for file in *.js; do
  if [ -f "$file" ]; then
    filename=$(basename "$file")
    # Check if file has the header
    if head -1 "$file" | grep -q "^/\* $filename Last Update"; then
      # Has header: check if file has been modified since last commit
      if git diff --quiet HEAD -- "$file" 2>/dev/null; then
        # No changes, skip
        continue
      else
        # Modified: use file mod time
        date=$(stat -f "%Sm" -t "%Y-%m-%d %H:%M" "$file")
      fi
    else
      # No header: add it, use the last commit time (excluding current if applicable) or mod time
      commit_dates=$(git log --format=%ci -- "$file" 2>/dev/null | head -2 | tail -1 | cut -d' ' -f1,2)
      if [ -z "$commit_dates" ]; then
        date=$(stat -f "%Sm" -t "%Y-%m-%d %H:%M" "$file")
      else
        date=$commit_dates
      fi
    fi
    update_header "$file" "$date"
    # Stage the updated file
    git add "$file"
  fi
done
