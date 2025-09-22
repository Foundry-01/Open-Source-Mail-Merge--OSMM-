## Task: Align project licensing to MIT and ensure consistency

### Goal
Adopt the MIT License for OSMM and make all project metadata, UI, and docs consistent.

### Steps
1. Decide license (MIT) — confirmed.
2. Add MIT `LICENSE` file at repo root.
3. Set `"license": "MIT"` in `website/package.json`.
4. Update footer in `website/src/pages/App.tsx` to say "Open-source under MIT".
5. Add a "License" section to `README.md` linking to the root `LICENSE` file.
6. Audit source for any remaining "Apache" mentions (exclude dependencies) and remove or correct.

### Current Next Step (with substeps)
- Implement changes:
  - Create root `LICENSE` (MIT)
  - Update `website/package.json` license field
  - Update footer text in `App.tsx`
  - Append License section to `README.md`
  - Search and clean any Apache mentions in source

### Running Note
- 2025-09-10: Created licensing nextsteps doc and outlined steps to standardize on MIT.


