# SIDH Marks Entry Automation

Automates entering student assessment marks on the **Skill India Digital Hub** portal.

## How it works

1. Script opens a browser
2. **You log in manually** (same browser, same session — no session expiry!)
3. Press ENTER in terminal
4. Script takes over and fills all 40 students' marks automatically

## Quick Start

### Test with one student first (dry run):
```bash
uv run python automate.py --dry-run --excel "Result Sheet _ 3391656, 20-05-2026.xlsx" --batch-id 3391656 --student CAN_37786291
```

### Run for all students:
```bash
uv run python automate.py --excel "Result Sheet _ 3391656, 20-05-2026.xlsx" --batch-id 3391656
```

## Options

| Flag | Description |
|------|-------------|
| `--excel "filename.xlsx"` | The name of the Excel file to read marks from |
| `--batch-id ID_NUMBER` | The Batch ID on the portal |
| `--dry-run` | Fill marks but don't submit |
| `--start-from N` | Resume from student #N |
| `--student CAN_XXX` | Process only one student |
| `--slow` | Extra delays for slow internet |

## Examples

```bash
# Run for a different Excel sheet and batch ID:
uv run python automate.py --excel "Result Sheet _ 3397371, 22-05-2026.xlsx" --batch-id 3397371

# Resume from student #15 (if it stopped at #14)
uv run python automate.py --excel "Result Sheet _ 3391656, 20-05-2026.xlsx" --batch-id 3391656 --start-from 15

# Slow internet
uv run python automate.py --excel "Result Sheet _ 3391656, 20-05-2026.xlsx" --batch-id 3391656 --slow
```

## Files

| File | Purpose |
|------|---------|
| `automate.py` | Main automation script |
| `gui.py` | Premium Desktop GUI wrapper |
| `build.bat` | One-click Windows standalone compiler |
| `automation.log` | Detailed log of what happened |
| `results.json` | Summary of results per student |
