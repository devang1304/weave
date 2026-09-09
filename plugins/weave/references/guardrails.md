# Guardrails: data handling, and the check before you send

Every skill reads this before touching source material, and shows the final
section to the user once a deliverable is built. Part 1 protects clients,
colleagues, and Netwoven's engagement terms while a skill is working; Part 2
is the last human checkpoint before a file leaves Netwoven.

## Contents

- Part 1: Data handling rules (refuse, warn, stop; labels; telemetry; reviewer)
- Part 2: Before you send (five checks)

---

# Part 1: Data handling rules for Weave skills

## Refuse, and say why

Do not put these into any generated file, spec, or telemetry event, and tell
the user what was skipped and why:

- Credentials, connection strings, API keys, tokens, passwords, MFA codes.
- Personal data beyond names and job roles: national IDs, dates of birth,
  home addresses, health details, salaries or performance notes about named
  people, bank details.
- Material a third party's NDA governs when the source says so, or content
  marked as another firm's confidential information.

## Warn once, then continue

- Internal candour about a client, a colleague, or a vendor. Leave it out of
  client files; say so under "I left out".
- Pricing, rates, or discounts that belong to a different client or
  engagement.
- Anything the source marks "do not share", "draft, not for client", or
  similar.
- Sources whose Purview label is Confidential (or the client's equivalent):
  proceed, and remind the user the output needs at least that label.

## Stop

- A client profile with `ai_permitted: false`: reply "This client's
  engagement terms restrict AI-assisted drafting. Ask the engagement lead."
  Produce no file and record no client hash.
- A source labelled Highly Confidential, Restricted, or encrypted: do not
  ingest it. Ask the user to supply an approved extract.

## Labels on outputs

Generated files keep the base template's Purview label unless a client
profile sets a higher default. Weave never lowers or removes a label. Part 2
below reminds the user to set the right label in Office.

## Telemetry

Events carry deliverable type, template version, skill, outcome, timing, and
a salted hash of the user and client. They never carry prompt text, file
names, paths, titles, or client names. Users can turn telemetry off in the
plugin settings or with `WEAVE_TELEMETRY=0`.

## Reviewer of record

Every generated file records "Draft prepared with Netwoven Weave; reviewed
by ____" in its revision history or notes. The reviewer's name is what makes
a deliverable Netwoven's; Weave use is expected, disclosed internally, and
never a mark against the author.

---

# Part 2: Before you send: five checks

Show this list once per deliverable, after the build, under the heading
"Check before sending". Keep the wording; shorten items that do not apply.

1. **Client name and date on the cover** read exactly as the client writes
   them (legal name on a SOW; short name is fine on a deck).
2. **Every number you did not supply** is traced to the source. Weave flags
   inferred or hedged figures; confirm or delete them.
3. **Nothing internal reached the file**: candour, other clients, internal
   system names, pricing from another engagement. The checks block template
   leakage; only you can judge substance.
4. **The sensitivity label** in Word or PowerPoint matches the content. SOWs
   and client documents usually need more than the template default.
5. **A named reviewer** has read it. Record the reviewer with
   `/weave:review` or in the revision history before it leaves Netwoven.
