# Recipe: adoption-training

Training deck on the presentation template: an end-user training session
for something being rolled out. This is one of three independent adoption
recipes (`adoption-strategy`, `adoption-comms`, `adoption-training`) -- pick
the one the request actually needs, not all three at once.

## Slide pattern

1. **Cover** (built automatically from `meta`).
2. **Objectives slide** (block/section id `objectives`, required): what
   attendees will be able to do after the session, stated as outcomes ("set
   up a Copilot prompt for a weekly status summary"), not as a list of
   features that will be shown.
3. **Agenda slide** (block/section id `agenda`, required): a simple
   time-boxed list -- topic and duration per line. Not a detailed run sheet;
   an attendee should be able to see at a glance what's covered and when.
4. **Key Topics slide(s)** (block/section id `key-topics`, required): the
   actual content, in plain, non-jargon language appropriate for end users.
   This is training material for people who may have never used the tool,
   not a technical spec -- write for that reader, not for IT.
5. **Resources slide** (block/section id `resources`, optional): links or
   names of supporting materials (quick reference card, recorded session,
   helpdesk contact). Drop the section if the source names none rather than
   inventing a resource that doesn't exist.
6. Closers (Confidentiality, Thank You, Closing) are added automatically.

## Asking for what's missing

Only `initiative` is required and usually isn't inferable from a generic
source -- ask what the session is about in one question if the source
doesn't make it obvious.

## Hard rules specific to this recipe

- Never invent a feature or capability the source material doesn't
  describe as real. If the source is vague about what a feature actually
  does, say less rather than guessing at functionality to fill the slide.
- Keep Key Topics in plain language throughout -- a training deck that
  reads like the product's technical documentation has the wrong register
  for its audience.
