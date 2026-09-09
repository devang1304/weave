---
name: ms-in-five
description: "Builds a short, cited PowerPoint deck (about five slides) that distills one Microsoft offering into a client-facing briefing: what it is, why it matters, how it works, and next steps. Every factual claim -- capability, price, licensing detail, release status -- traces to a page fetched live via WebFetch from learn.microsoft.com or www.microsoft.com, cited on its slide with the exact URL and the date retrieved; nothing is stated from the model's own training knowledge, and anything that can't be verified is dropped or flagged rather than guessed. Use when the user asks to explain a Microsoft product simply, distill a Microsoft offering for a client, or brief a client audience on a Microsoft capability. Not for a full assessment, a sales proposal, an internal design document, or a deck covering more than one offering."
disable-model-invocation: true
argument-hint: "[Microsoft offering name]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *), WebFetch(domain:learn.microsoft.com), WebFetch(domain:www.microsoft.com)
---

# Microsoft offering in five slides

Recipe: `${CLAUDE_PLUGIN_ROOT}/recipes/ms-in-five.md` (deck, `ms-in-five`) -- read it
before writing the spec. Its whole point is citation discipline: nothing
about the offering reaches a slide unless it was just read from Microsoft's
own documentation.

## 1. Check the environment

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

## 2. Get the offering

Use the `offering` argument if given; otherwise ask the recipe's one
question (`inputs[].ask`): which Microsoft offering or product to distill.

## 3. Fetch before writing a single word of content

Before drafting any slide, fetch 1 to 3 authoritative pages via WebFetch:
`learn.microsoft.com` first, for capability and how-it-works detail;
`www.microsoft.com` for positioning, pricing, and licensing framing. For
every page fetched, record the exact URL and today's date -- this is the
citation each slide's `source` field will carry. No other domain is
reachable; if the two together do not settle a claim, it goes to the flag
list as unverified, never onto a slide from memory.

## 4. Write the spec and build

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write the spec per `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md` with
`weave.recipe: "ms-in-five"`: four required sections (`what-it-is`,
`why-it-matters`, `how-it-works`, `next-steps`), each content slide's
`source` field set to "Source: <url>, retrieved <date>". Follow
`${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` for register. Five slides
total -- cover plus the four sections; no sixth content slide.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --recipe ms-in-five --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT --quick --json
```

Fix every hard failure and high-severity finding, then rebuild.

## 5. Edit or rebuild

`weave.py spec-from-file FILE --out spec.json`, edit, rebuild with
`--replace`. Re-verify any edited fact against a freshly fetched source
before rebuilding -- an edit is not exempt from the citation rule.

## 6. Deliver

**Check before sending** · **I assumed** · **I left out** · **Open
decisions** -- name anything that could not be verified via WebFetch here --
then `Saved to: <absolute path>`.
