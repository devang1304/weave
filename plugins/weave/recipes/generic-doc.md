# Recipe: generic-doc

The default recipe for any Word document without a more specific recipe.
It constrains nothing but the pattern below; Phase 2 recipes (assessment,
status report, proposal, runbook) add required sections and validators.

## Section pattern

1. **Executive summary** (level 1): governing thought first sentence, then
   SCQA in one or two paragraphs, then the key line as 3 to 5 bullets that
   mirror the level-1 headings that follow. A reader who stops here has the
   answer.
2. **One level-1 heading per key-line statement**, phrased as the insight
   ("Permissions sprawl blocks the Copilot rollout", not "Permissions").
   Recurse with levels 2 to 4; never skip a level; no heading with exactly
   one child.
3. **Evidence** as tables (Netwoven Table 1 for comparisons, Table 2 for
   plain data, Table 3 for callouts and single-column lists) or figures at
   6.5 inches, each with a caption that states the takeaway. Captions
   auto-number; never type "Table 1".
4. **Recommendations or next steps** with owner and date where the source
   gives them; `[TO BE CONFIRMED]` where it does not.
5. **Appendix** only for material the reader may need but the argument does
   not (raw inventories, glossaries, method notes).

## Internal variant

`meta.internal: true`. Confidentiality page removed; Company is Netwoven.
The governing thought is usually a decision needed, a change announced, or
a procedure to follow; name it in the first lines.

## Spec skeleton

```json
{"weave": {"spec_version": "2.0", "kind": "doc", "recipe": "generic-doc",
           "base": "NW_Document_Base_2026.docx", "revision": 1},
 "meta": {"title": "...", "client": "...", "date": "September 2026", "internal": false},
 "blocks": [
   {"id": "b1", "type": "heading", "level": 1, "text": "Executive summary"},
   {"id": "b2", "type": "para", "text": "..."},
   {"id": "b3", "type": "bullets", "items": ["...", "..."]},
   {"id": "b4", "type": "heading", "level": 1, "text": "<Key-line insight 1>"}
 ]}
```
