# Pre-Commit Hook for Versioning JS Files

This script automatically adds or updates headers in root-level `.js` files with the filename, last update date/time, and SHA256 hash. This helps track versions in container-attached Google Apps Script (GAS) projects.

## Header Format

```text
/* Filename.js Last Update YYYY-MM-DD HH:MM <SHA256>
```

## How It Works

- **For files with headers**: Updates the date/hash if the file has been modified since the last commit.
- **For files without headers**: Adds a header using the second-to-last commit date (to avoid using the current commit time).
- Computes SHA256 from line 2 onward to prevent the hash from changing the hash.

## Setup

1. Copy `pre-commit-hook.sh` to your project's `scripts/` folder.

2. Create a symbolic link or copy to `.git/hooks/pre-commit`:

   ```bash
   ln -s ../../scripts/pre-commit-hook.sh .git/hooks/pre-commit
   ```

   Or copy the file:

   ```bash
   cp scripts/pre-commit-hook.sh .git/hooks/pre-commit
   chmod +x .git/hooks/pre-commit
   ```

3. Ensure the hook is executable: `chmod +x .git/hooks/pre-commit`

## Usage

- The hook runs automatically on `git commit`.
- To test manually: `./scripts/pre-commit-hook.sh`
- Skip the hook for a commit: `git commit --no-verify`

## Notes

- Only affects root-level `.js` files.
- Uses Git history for dates to ensure accuracy.
- Safe for shared repositories; others can set it up similarly.
