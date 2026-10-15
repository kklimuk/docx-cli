# Review: contract-finalize (bucket: editing, review finalize), arm: Anthropic docx skill (haiku), r2

## Verdict
| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true** (3 output pages rendered) |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** |

All 13 rubric conditions pass.

## Rubric check
1. **Net 30 accepted.** p7 reads "All payments are Net 30 from the Company's acceptance" as clean text. The baseline's "Deleted: Net 90 from the Company" balloon is gone from the render.
2. **One (1) year accepted.** p18 reads "For one (1) year after the Agreement ends" as clean text, with no balloon.
3. **Liability cap rejected.** p24 reads "one hundred dollars ($100)". The "$100,000" insertion is gone.
4. **§11 Personal Guarantee deletion rejected.** The heading "11. Personal Guarantee." and its body paragraph are restored. Later sections renumber to 12–15, which is correct with §11 back.
5. **Indemnification spacing accepted.** p22 has `line-spacing="1.5"` as live formatting. `document.xml` has 0 `pPrChange` elements.
6. **No other pending changes.** `docx track-changes list` returns "No tracked changes", and `document.xml` has 0 `w:ins`, `w:del`, `pPrChange` or `rPrChange` elements.
7. **Every comment has a reply.** `comments list` shows replies c4–c7, with `parentId` c0, c1, c2 and c3.
8. **Replies are anchored.** Every reply has a non-null `parentId` and a non-empty `anchor.startBlockId` (p3, p12, p22, p30).
9. **Replies are visible in the read view.** `read --comments` shows `[^c4]`–`[^c7]`, each with a `↳ cN` marker.
10. **Auto-renewal resolved.** c0 is `resolved: true`. The render shows its NF2 reply greyed out, which is how Word shows a resolved thread.
11. **Arbitration resolved.** c3 is `resolved: true`, and the render shows NF8 greyed out.
12. **IP open.** c1 is `resolved: false`.
13. **Indemnification open.** c2 is `resolved: false`.

`commentsExtended.xml` is well-formed. Each reply has a `w15:paraIdParent` pointing to its parent's paraId, and `done` is set to 1 on the c0 and c3 threads only.

## Merits
- Each tracked change got the right accept or reject. This includes the hard case: rejecting a deleted paragraph mark to restore §11, which the agent did by hand in XML.
- The pPrChange was accepted correctly. The 1.5 spacing stays, and the change record is gone.
- Replies are properly threaded in both `commentsExtended` (`paraIdParent`) and `document.xml` (nested range markers). Word renders them as threaded replies under each parent.
- Only the resolved threads are marked resolved. IP and indemnification stay open.
- Comparing output and baseline, nothing else changed. Header, footer, table (including its own "Net 90 after final acceptance" cell, which was never a tracked change), signature block and numbering are all intact. Both renders are 3 pages.

## Demerits
- **minor: the replies are boilerplate.** All four replies say "Declined — standard term.", including on the two threads the task describes as "fully addressed" and resolved. Replying "Declined" and then resolving the thread reads oddly as a negotiating position. The task allowed short positions, but these are copied from the example rather than written for each comment.
- **minor (hygiene): comments.xml was re-serialized.** `comment.py` rewrote it through minidom, so the new `w:comment` elements contain pretty-printed whitespace between elements. Word ignores it and the render is clean. docx-cli's `comments list` shows it as leading and trailing newlines in the reply `text`. That is a quirk of docx-cli's reader, not a defect in the competitor's output.

## Frictions (agent-reported)
- **Blocker:** the skill has no per-revision accept/reject. `accept_changes.py` accepts everything, so the agent edited `document.xml` by hand for every revision, including the deleted paragraph mark and the pPrChange.
- The skill has no way to resolve a comment. The agent had to create the `w15:done` and `commentsExtended` parent entries by hand.
- `comment.py --parent` needs the reply's range markers nested by hand before the reply shows up.
- `validate.py --author` reports rejected revisions as mismatches, which is misleading.
- Running `comment.py` with `python3 -I` fails because the `office` module can't be found.
- Pre-existing paraId warnings blocked a clean validator pass.
- Listing the revisions took an ad-hoc regex script.

**Read:** the output is correct, but getting there was a slog. A weak agent had to hand-edit OOXML revision markup and comment-extension parts because the skill has no tools for per-change accept/reject, reply threading or resolution. It succeeded this time, but the path is fragile.
