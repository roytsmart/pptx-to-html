# pptx-to-html

Turn a PowerPoint deck into a single web page, with its movies playing in
place instead of frozen on their first frame.

Every slide keeps its layout. Pictures, text, and shapes sit where PowerPoint
puts them, as shares of the slide, so a slide scales as a whole. The page
lists the slides down the screen to read, and has a button (or the `f` key)
to present them full screen, one at a time.

## Installing

```bash
pip install git+https://github.com/roytsmart/pptx-to-html
```

## Using it

```bash
pptx-to-html talk.pptx docs/ --subtitle "SPD 2026 · the slides as they were presented"
```

This writes `docs/index.html` and copies the deck's pictures and movies into
`docs/media/`. Publish the folder with GitHub Pages, or open the page
directly.

Movies embedded in a deck are often much larger than they need to be on the
web. `--reencode` re-encodes them as H.264 with `ffmpeg`, at most 1920 pixels
wide by default:

```bash
pptx-to-html talk.pptx docs/ --reencode
```

`--media` swaps a file embedded in the deck for another address, by the name
the deck gives it (the names in `docs/media/`, or the file name of a movie the
deck only links to). Point a picture at an `.mp4` and it plays as a movie,
which is how to replace a heavy animated GIF with the movie it was made from:

```bash
pptx-to-html talk.pptx docs/ \
    --media media3.mp4=../figures/inversion.mp4 \
    --media image1.gif=../figures/cinemagraph.mp4
```

The same is available from Python:

```python
import pptx_to_html

pptx_to_html.convert("talk.pptx", "docs", reencode=pptx_to_html.Reencode())
```

## What it handles

- Pictures, including cropped ones and SVGs.
- Movies, with their poster frames.
- Text in placeholders and text boxes, with the formatting it inherits from
  the layout, the master, and the theme: sizes, colors, typefaces, bold and
  italic, alignment, line spacing, space between paragraphs, autofit, and
  bulleted and numbered lists.
- Filled and outlined shapes (rectangles, rounded rectangles, ellipses,
  octagons, triangles, diamonds), and lines with arrowheads.
- Grouped shapes, slide backgrounds, and shapes drawn on layouts and masters.

Tables, charts, SmartArt, and sounds are left out, with a warning naming the
slide. Animations are shown in their final state. Text is set in the typeface
the theme asks for when the reader has it, and is shrunk to fit its box when a
substitute sets wider.
