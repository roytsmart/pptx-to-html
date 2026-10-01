"""The command-line interface: ``pptx-to-html deck.pptx output/``."""

import argparse
import shutil
import sys
import warnings

from ._convert import ConversionWarning, Reencode, convert


def main(argv: "list[str] | None" = None) -> int:
    parser = argparse.ArgumentParser(
        prog="pptx-to-html",
        description="Turn a PowerPoint deck into a single web page.",
    )
    parser.add_argument("deck", help="the .pptx file to convert")
    parser.add_argument("output", help="the folder to write index.html and media/ into")
    parser.add_argument("--title", help="the heading of the page (default: the deck's title)")
    parser.add_argument("--subtitle", help="a line under the heading")
    parser.add_argument(
        "--media",
        action="append",
        default=[],
        metavar="NAME=ADDRESS",
        help=(
            "use ADDRESS instead of the copy of NAME embedded in the deck, "
            "e.g. media1.mp4=../figures/movie.mp4 (repeatable)"
        ),
    )
    parser.add_argument(
        "--reencode",
        action="store_true",
        help="re-encode movies as H.264 with ffmpeg to make them smaller",
    )
    parser.add_argument("--ffmpeg", default=None, help="the ffmpeg executable (default: on PATH)")
    parser.add_argument("--crf", type=int, default=26, help="H.264 quality, higher is smaller (default: 26)")
    parser.add_argument("--max-width", type=int, default=1920, help="widest a movie may be (default: 1920)")
    parser.add_argument("--include-hidden", action="store_true", help="include hidden slides")
    args = parser.parse_args(argv)

    media = {}
    for item in args.media:
        name, separator, address = item.partition("=")
        if not separator:
            parser.error(f"--media expects NAME=ADDRESS, not {item!r}")
        media[name] = address

    reencode = None
    if args.reencode:
        ffmpeg = args.ffmpeg or shutil.which("ffmpeg")
        if ffmpeg is None:
            parser.error("--reencode needs ffmpeg; put it on PATH or pass --ffmpeg")
        reencode = Reencode(ffmpeg=ffmpeg, crf=args.crf, max_width=args.max_width)

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always", ConversionWarning)
        index = convert(
            args.deck,
            args.output,
            title=args.title,
            subtitle=args.subtitle,
            media=media,
            reencode=reencode,
            include_hidden=args.include_hidden,
        )
    for warning in caught:
        print(f"warning: {warning.message}", file=sys.stderr)
    print(index)
    return 0


if __name__ == "__main__":
    sys.exit(main())
