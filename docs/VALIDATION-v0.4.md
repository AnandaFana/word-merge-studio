# v0.4 validation

## Automated checks

The public regression suite covers comparison, independent formatting, chapter append and section assembly. Personal document fixtures are opt-in using `RUN_LOCAL_SAMPLES=1`; public CI never requires or uploads them.

2026-09-12 local results: public suite **51 passed, 3 optional tests skipped**; opt-in local suite **54 passed** with a 60-second per-test limit. The first local-sample run hit the default 5-second timeout; the longer run completed successfully. TypeScript checking and production build passed. Language switching is tested to leave document text and section matching unchanged.

Section-specific checks cover:

- Normalised title matching and complete, non-overlapping block partitioning.
- Reordered groups and sources in both original and unified formatting modes.
- Text-bearing content retained except explicitly replaced root headings and excluded cards.
- Three navigation levels, table preservation, image byte identity and OPC relationship resolution.
- Excluded-card image removal, front matter, documents without matching headings, duplicate/empty plan rejection and source immutability.
- Source font inheritance in preserve mode and unified paragraph styles in format mode.

## Browser checks

Checked locally with the in-app browser on 2026-09-12: English/Chinese switching, fictional demo loading, group movement, title editing, card exclusion and undo, preview in original/unified modes, three loaded images and three tables in the result.

## Evidence limits

Browser preview uses `docx-preview`, not Microsoft Word. A native render attempt failed because the managed environment has no `soffice.exe`. Native Word pagination and full layout fidelity have **not** been verified. Final user documents, complex tables, floating drawings and fonts still require a Word review.

The tests are local package/behaviour checks, not a claim that arbitrary Word documents are supported. See [section guide](SECTIONS.md) for rejected structures and size limits.
