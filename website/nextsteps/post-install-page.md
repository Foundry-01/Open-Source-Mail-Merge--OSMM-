# Post Install Page

Goal: Add a new website page "Post Install" with clear instructions on how to use OSMM after installing the add‑on. Match existing site styles and navigation.

## Current Step: Implement page and route

### Plan
- Create React page `src/pages/PostInstall.tsx` using existing `DocPage` layout for consistent styling
- Add route `/post-install` in `src/pages/App.tsx`
- Add header nav link "Post Install" pointing to `/post-install`
- Build and preview to validate

### Subtasks
- [ ] Create `PostInstall.tsx` with sections and TOC
- [ ] Register route in `App.tsx`
- [ ] Add nav link in `App.tsx`
- [ ] Build and verify locally
- [ ] Update `website/cmds.md` with verification snippet

## Running Note
- Created this nextsteps document to track work on the Post Install page
- Implemented `PostInstall.tsx` using `DocPage` for consistent styling
- Added `/post-install` route and nav link; built successfully
- Updated copy per feedback: simplified Email Address instruction with sub‑bullets; removed "(no BCC)" from send step; added BCC alongside CC tips

## Running Note
- Created this nextsteps document to track work on the Post Install page


