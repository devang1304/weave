#!/usr/bin/env python3
"""
new_deck_pptx.py - create, finalize and demo Netwoven decks on
NW_Presentation_Base_2026.pptx.

The base ships six slides: an instructional slide, an empty demo cover, an
empty demo agenda, a "SECTION DIVIDER" demo, the branded "THANK YOU" closer
(Blank layout, 12 SVG ribbons) and the contact closer ("Closing Slide" layout).
Only the two closers are worth keeping. python-pptx has no delete-slide API and
appends new slides at the end, so this script owns the sldIdLst surgery.

Usage
  # 1) shell: cover + confidentiality + kept closers; demo slides gone
  python new_deck_pptx.py create BASE OUT --title "Deck title" --subtitle "Client / date" \
      --client "Contoso" [--cover-image a.jpg[,b.jpg]] [--internal] \
      [--no-confidentiality] [--no-thankyou] [--no-closing]

  # 2) add content slides to OUT with nw_pptx_helpers (they land after the closers)

  # 3) closers back to the tail, empty placeholders removed, app.xml fixed
  python new_deck_pptx.py finalize OUT [--title "Deck title"] [--prune-empty]

  # sample deck that exercises every helper (used by tools/smoke_test.py)
  python new_deck_pptx.py demo BASE OUT [--client "Contoso"] [--internal]

--internal (or --no-confidentiality) skips the Statement of Confidentiality slide;
otherwise --client is required because the statement names the client.

Requires: python-pptx, lxml, and nw_pptx_helpers.py next to this file.
"""
import argparse
import datetime as dt
import os
import random
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pptx  # noqa: E402
from pptx.util import Inches  # noqa: E402

import nw_pptx_helpers as H  # noqa: E402

COVER_IMAGES_DIR = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "cover-images"))


def pick_cover_images(n=2):
    """Randomly pick up to `n` images from shared/assets/cover-images/ for the
    cover's portrait slots. [] when the pool doesn't exist or is empty --
    add_cover() already leaves an unfilled slot cleanly removed, so a caller
    doesn't need to special-case an empty pool."""
    if not os.path.isdir(COVER_IMAGES_DIR):
        return []
    pool = [f for f in os.listdir(COVER_IMAGES_DIR)
            if f.lower().endswith((".jpg", ".jpeg", ".png")) and not f.startswith(".")]
    random.shuffle(pool)
    return [os.path.join(COVER_IMAGES_DIR, f) for f in pool[:n]]


# ---------------------------------------------------------------------------
# create
# ---------------------------------------------------------------------------

def _drop_reason(slide):
    """Why a base slide is dropped (None = keep)."""
    text = H.slide_text(slide)
    low = text.lower()
    for m in H.DROP_MARKERS:
        if m.lower() in low:
            return "marker %r" % m
    lname = H.layout_name(slide)
    if lname in H.BROKEN_LAYOUTS:
        return "broken layout %r" % lname
    if lname == "Title Slide for Verticals":
        return "demo cover (empty title)"
    if lname == "Agenda":
        return "demo agenda"
    return None


def build_shell(prs, title, subtitle=None, client=None, cover_images=None, internal=False,
                no_confidentiality=False, no_thankyou=False, no_closing=False, log=print):
    """Turn the freshly opened base into [cover, confidentiality?, thank you?, closing?]."""
    want_conf = not (internal or no_confidentiality)
    if want_conf and not client:
        raise SystemExit("--client is required unless --internal or --no-confidentiality is given")

    closers = H.tag_closers(prs)
    dropped = []
    for slide in list(prs.slides):
        kind = H.closer_kind(slide)
        if kind == H.TAG_THANKYOU and not no_thankyou:
            continue
        if kind == H.TAG_CLOSING and not no_closing:
            continue
        reason = _drop_reason(slide) or (
            "closer disabled" if kind else "not a recognised closer")
        dropped.append((H.slide_title_text(slide) or H.layout_name(slide), reason))
        H.delete_slide(prs, slide)
    # surviving closers are slide5/slide6 in the base: renumber so new slides never collide
    H.renumber_slide_parts(prs)

    cover = H.add_cover(prs, title, subtitle, cover_images)
    H.move_slide(prs, cover, 0)
    if want_conf:
        conf = H.add_confidentiality(prs, client, dt.date.today().year)
        H.move_slide(prs, conf, 1)
    H.move_closers_last(prs)
    H.remove_dangling_section_ids(prs)
    H.renumber_slide_parts(prs)  # saved shell is slide1..N, so later add_slide calls are safe
    prs.core_properties.title = title
    if client and not internal:
        prs.core_properties.subject = client
    prs.core_properties.author = "Netwoven"
    prs.core_properties.last_modified_by = "Netwoven"
    for name, reason in dropped:
        log("  dropped %-40s %s" % (repr(name[:38]), reason))
    kept = [H.closer_kind(s) for s in prs.slides if H.closer_kind(s)]
    log("  kept closers: %s" % (kept or "none"))
    return prs


def _save(prs, out):
    # tmp + move, not prs.save(out) directly -- see new_deliverable_docx.py's
    # finish_document() for why.
    tmp = out + ".tmp"
    prs.save(tmp)
    shutil.move(tmp, out)
    n = H.patch_app_xml(out)
    return n


def print_order(prs):
    print("Final slide order:")
    for i, s in enumerate(prs.slides, start=1):
        tag = H.closer_kind(s)
        print("  %2d. %-50s [%s%s]" % (i, H.slide_title_text(s)[:50] or "(untitled)",
                                       H.layout_name(s), " / " + tag if tag else ""))


def cmd_create(args):
    prs = pptx.Presentation(args.base)
    images = [p.strip() for p in (args.cover_image or "").split(",") if p.strip()]
    for p in images:
        if not os.path.exists(p):
            raise SystemExit("cover image not found: %s" % p)
    if len(images) > 2:
        raise SystemExit("--cover-image takes at most two files")
    print("Creating shell from %s" % args.base)
    build_shell(prs, args.title, args.subtitle, args.client, images, args.internal,
                args.no_confidentiality, args.no_thankyou, args.no_closing)
    n = _save(prs, args.out)
    print("OK: %s written (%d slides). Add content slides with nw_pptx_helpers, then run:\n"
          "    new_deck_pptx.py finalize %s --title %r" % (args.out, n, args.out, args.title))
    print_order(pptx.Presentation(args.out))


# ---------------------------------------------------------------------------
# finalize
# ---------------------------------------------------------------------------

def finalize_deck(path, title=None, prune_empty=True, log=print):
    prs = pptx.Presentation(path)
    moved = H.move_closers_last(prs)
    removed = 0
    if prune_empty:
        for s in prs.slides:
            removed += H.remove_empty_placeholders(s)
    dangling = H.remove_dangling_section_ids(prs)
    H.renumber_slide_parts(prs)
    if title:
        prs.core_properties.title = title
    elif not (prs.core_properties.title or "").strip():
        first = list(prs.slides)[0] if len(prs.slides) else None
        if first is not None:
            prs.core_properties.title = H.slide_title_text(first)
    prs.core_properties.last_modified_by = "Netwoven"
    n = _save(prs, path)
    log("OK: finalized %s (%d slides; closers moved: %s; empty placeholders removed: %d; "
        "dangling section ids removed: %d)" % (path, n, moved or "none", removed, dangling))
    return pptx.Presentation(path)


def cmd_finalize(args):
    prs = finalize_deck(args.deck, args.title, prune_empty=not args.no_prune_empty)
    print_order(prs)


# ---------------------------------------------------------------------------
# demo - exercises every helper
# ---------------------------------------------------------------------------

def build_demo_content(prs):
    """Append the sample body slides (call after build_shell, before finalize)."""
    C = H.TITLE_ONLY_CONTENT

    H.add_section_header(prs, "Where the programme stands today", "Findings from the discovery phase")

    H.add_title_content(prs, "Five findings explain the slow adoption of the new intranet", [
        "Search returns stale documents because 40% of sites were never migrated",
        "Site owners lack a governance playbook, so permissions drift within weeks",
        "Two regional hubs duplicate the global news feed",
        "Mobile traffic is 31% of visits but the home page is not responsive",
        "No adoption metrics are reviewed by leadership",
    ])

    H.add_two_content(prs, "The current estate and the target estate differ on three axes",
                      ["Current state", "312 sites, 40% unmigrated", "Permissions managed ad hoc",
                       "No published governance"],
                      ["Target state", "180 hub-connected sites", "Role-based access review each quarter",
                       "Governance playbook owned by IT"],
                      left_levels=[0, 1, 1, 1], right_levels=[0, 1, 1, 1])

    H.add_three_column(prs, "Three workstreams deliver the target state in nine months",
                       [["Inventory all 312 sites, then migrate the 180 keepers in three waves"],
                        ["Publish the governance playbook and train every site owner"],
                        ["Tune search and give leadership a monthly usage dashboard"]],
                       headings=["Consolidate", "Govern", "Measure"])

    H.add_three_stat(prs, [("312", "sites in the current estate"),
                           ("40%", "of sites never migrated"),
                           ("9 mo", "to reach the target state")])

    s = H.add_title_only(prs, "Wave 1 moves the eight highest-traffic sites first")
    H.add_table(s, ["Wave", "Sites", "Monthly visits", "Owner"],
                [["1", "8", "184,000", "Corporate Comms"],
                 ["2", "42", "96,500", "Regional IT"],
                 ["3", "130", "41,200", "Business units"],
                 ["Retire", "132", "6,300", "IT"]],
                C["left"], C["top"], C["width"], Inches(2.4), col_widths=[1, 1, 1.4, 2], font_pt=12)
    H.add_source_line(s, "Source: SharePoint usage report, July 2026; site inventory, August 2026")

    s = H.add_title_only(prs, "Visits recover within a quarter of each migration wave")
    H.add_chart(s, "column", ["Q1", "Q2", "Q3", "Q4"],
                {"Wave 1 sites": [120, 150, 178, 184],
                 "Wave 2 sites": [80, 82, 90, 97],
                 "Wave 3 sites": [40, 38, 39, 41]},
                C["left"], C["top"], C["width"], Inches(4.6), data_labels=False, number_format="#,##0")
    H.add_source_line(s, "Monthly visits in thousands. Source: SharePoint usage report, 2026")

    s = H.add_title_only(prs, "Retired sites hold less than 3% of total traffic")
    H.add_chart(s, "pie", ["Wave 1", "Wave 2", "Wave 3", "Retire"],
                {"Share of visits": [56, 29, 12, 3]},
                Inches(3.4), C["top"], Inches(6.5), Inches(4.6), data_labels=True, number_format="0%")
    H.add_source_line(s, "Share of monthly visits by wave. Source: SharePoint usage report, July 2026")

    H.add_comparison(prs, "A phased migration beats a big-bang cutover on risk and cost",
                     "Phased (recommended)",
                     ["Three waves over nine months", "Rollback per wave", "Change fatigue spread out"],
                     "Big bang",
                     ["Single weekend cutover", "No rollback once search re-indexes",
                      "Helpdesk peak of 4x normal volume"])

    H.add_big_statement(prs, "Consolidating to 180 sites cuts search noise in half",
                        eyebrow="THE HEADLINE")
    return prs


def cmd_demo(args):
    prs = pptx.Presentation(args.base)
    client = args.client or ("Netwoven" if args.internal else "Contoso")
    print("Building demo deck from %s" % args.base)
    build_shell(prs, "Intranet consolidation: findings and roadmap",
                "%s / %s" % (client, dt.date.today().strftime("%B %Y")),
                client, None, args.internal, False, False, False)
    build_demo_content(prs)
    _save(prs, args.out)
    print("OK: demo content written to %s; run finalize next." % args.out)
    if not args.no_finalize:
        prs = finalize_deck(args.out, "Intranet consolidation: findings and roadmap")
        print_order(prs)


# ---------------------------------------------------------------------------

def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    c = sub.add_parser("create", help="cover + confidentiality + closers shell")
    c.add_argument("base")
    c.add_argument("out")
    c.add_argument("--title", required=True)
    c.add_argument("--subtitle", default=None)
    c.add_argument("--client", default=None)
    c.add_argument("--cover-image", default=None, help="one or two image files, comma separated")
    c.add_argument("--internal", action="store_true", help="internal deck: no confidentiality slide")
    c.add_argument("--no-confidentiality", action="store_true")
    c.add_argument("--no-thankyou", action="store_true")
    c.add_argument("--no-closing", action="store_true")
    c.set_defaults(func=cmd_create)

    f = sub.add_parser("finalize", help="closers last, empty placeholders removed, app.xml fixed")
    f.add_argument("deck")
    f.add_argument("--title", default=None)
    f.add_argument("--prune-empty", action="store_true", default=True,
                   help="remove unfilled placeholders (default on)")
    f.add_argument("--no-prune-empty", action="store_true", help="keep unfilled placeholders")
    f.set_defaults(func=cmd_finalize)

    d = sub.add_parser("demo", help="sample deck exercising every helper (finalized)")
    d.add_argument("base")
    d.add_argument("out")
    d.add_argument("--client", default=None)
    d.add_argument("--internal", action="store_true")
    d.add_argument("--no-finalize", action="store_true", help="stop before finalize (for testing)")
    d.set_defaults(func=cmd_demo)

    args = ap.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
