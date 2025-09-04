## Home hero copy refresh

**Goal**: Emphasize that OSMM is free, easy to use, and locally hosted (runs in the user's Google account). Call out primary audiences: clubs, nonprofits, and small businesses that avoid paid mail‑merge tools.

### Current step
- Update hero subtext in `src/pages/Home.tsx` accordingly.

### Substeps for current step
- Edit `src/pages/Home.tsx` hero paragraph.
- Run lints and fix any TypeScript issues.
- Sanity-check page locally.

### Running notes
- Edited hero paragraph to highlight free, easy‑to‑use, locally hosted value props and target audiences.
- Fixed TS return type in `Home` component to `React.ReactElement` to satisfy linter.
- Changed hero to single `h1` with `<br />` for second line: "No BCC, no problem".
- Updated pronunciation easter egg to a clearer aside: "(/ aw·səm /)" with same tooltip.

### Next step
- If we add more positioning, consider aligning bullets and CTA text with the updated messaging.


