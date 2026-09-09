# Cover image pool

Deck covers have two portrait picture slots on the "Title Slide for
Verticals" layout. When a spec doesn't set `meta.cover_images` explicitly,
`build_deck()` picks up to 2 images at random from this folder
(`new_deck_pptx.pick_cover_images()`). An empty folder is a no-op -- the
slots stay unfilled, same as today.

Drop `.jpg`/`.jpeg`/`.png` files directly in this folder to add them to the
pool. No manifest or naming convention required; every image here is
eligible to be picked for every deck.

Guidance for what to add (portrait orientation, since the slots are tall
and narrow):
- Real Netwoven people, offices, or work in progress -- not generic stock.
- Nothing client-identifying (no visible client names, logos, or screens).
- A few different subjects so repeat decks don't always show the same photo.

To force specific images on one deck instead of a random pick, set
`meta.cover_images` in the spec to a list of 0-2 absolute paths.
