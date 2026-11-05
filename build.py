# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx", "python-docx", "pillow"]
# ///
"""Builds about_me.pptx and talk_script.docx.

Drop photos into photos/ named after the keys in IMAGES (e.g. photos/running.jpg)
and re-run `uv run build.py` – they replace the placeholders automatically.
"""

from pathlib import Path
import base64
import io
import subprocess

from docx import Document
from docx.shared import Cm, Pt as DocPt

from PIL import Image, ImageDraw, ImageFont, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

HERE = Path(__file__).parent
PHOTOS = HERE / "photos"
CACHE = HERE / "build"
SITE = HERE / "docs"  # GitHub Pages serves this folder

NAVY = RGBColor(0x0B, 0x16, 0x3A)
TILE = RGBColor(0x16, 0x25, 0x55)
RED = RGBColor(0xE3, 0x1B, 0x3B)
YELLOW = RGBColor(0xFF, 0xC9, 0x06)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY = RGBColor(0xAE, 0xB8, 0xD6)
FONT = "Arial"
RADIUS = 0.14  # corner radius (inches) shared by tiles and pictures

# key -> placeholder label
IMAGES = {
    "running": "PHOTO\nhalf marathon",
    "hiking": "PHOTO\nhiking",
    "skiing": "PHOTO\nskiing",
    "accenture_logo": "LOGO\nAccenture",
    "connells_logo": "LOGO\nConnells",
    "redbull_car": "PHOTO\nRed Bull car",
    "album_dire_straits": "COVER\nMaking\nMovies",
    "album_arctic_monkeys": "COVER\nWhatever\nPeople Say…",
    "album_metallica": "COVER\nRide the\nLightning",
}


# ---------- slide helpers (all positions in inches) ----------


def _font(size):
    try:
        return ImageFont.truetype("/Library/Fonts/Arial Unicode.ttf", size)
    except OSError:
        return ImageFont.load_default(size)


def _placeholder(key, w, h):
    size = (int(w * 150), int(h * 150))
    img = Image.new("RGB", size, (0x2A, 0x3A, 0x6E))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([4, 4, size[0] - 5, size[1] - 5], radius=RADIUS * 150, outline=(0x5A, 0x6A, 0x9E), width=4)
    fs = int(min(size) * 0.16)
    while (
        fs > 8
        and d.multiline_textbbox((0, 0), IMAGES[key], font=_font(fs))[2]
        > size[0] * 0.85
    ):
        fs -= 1
    d.multiline_text(
        (size[0] / 2, size[1] / 2),
        IMAGES[key],
        fill=(0xC8, 0xD0, 0xE8),
        font=_font(fs),
        anchor="mm",
        align="center",
    )
    return img


def _photo(src, key, w, h):
    img = ImageOps.exif_transpose(Image.open(src)).convert("RGBA")
    size = (int(w * 200), int(h * 200))
    if key.endswith("_logo"):  # keep the whole logo, on white
        canvas = Image.new("RGBA", size, "white")
        img = ImageOps.contain(img, (int(size[0] * 0.85), int(size[1] * 0.8)))
        canvas.alpha_composite(
            img, ((size[0] - img.width) // 2, (size[1] - img.height) // 2)
        )
        img = canvas
    else:  # crop to fill the box
        img = ImageOps.fit(img, size)
    return img


def image(slide, key, x, y, w, h):
    src = next(
        (
            p
            for p in PHOTOS.glob(f"{key}.*")
            if p.suffix.lower() in {".jpg", ".jpeg", ".png", ".webp"}
        ),
        None,
    )
    img = (_photo(src, key, w, h) if src else _placeholder(key, w, h)).convert("RGBA")
    # corners are baked into the image so every viewer shows them
    mask = Image.new("L", img.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, img.width - 1, img.height - 1], radius=RADIUS * img.width / w, fill=255)
    img.putalpha(mask)
    path = CACHE / f"{key}.png"
    img.save(path)
    slide.shapes.add_picture(str(path), Inches(x), Inches(y), Inches(w), Inches(h))


def text(slide, x, y, w, h, lines, align=PP_ALIGN.LEFT):
    """lines: list of (text, size, colour, bold, space_before_pt)."""
    tf = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, (s, size, colour, bold, before) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(before)
        r = p.add_run()
        r.text = s
        r.font.name, r.font.size, r.font.color.rgb, r.font.bold = (
            FONT,
            Pt(size),
            colour,
            bold,
        )


def rect(slide, shape, x, y, w, h, colour):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = colour
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def tile(slide, x, y, w, h, label):
    rect(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, TILE).adjustments[0] = RADIUS / min(w, h)
    rect(slide, MSO_SHAPE.RECTANGLE, x + 0.25, y + 0.2, 0.07, 0.28, RED)
    text(slide, x + 0.42, y + 0.2, w - 0.6, 0.3, [(label, 13, YELLOW, True, 0)])


def gallery(slide, x, y, w, items, img_h, square=False, gap=0.2):
    """Row of images with a bold title and a grey caption under each."""
    cell = (w - gap * (len(items) - 1)) / len(items)
    img_w = min(cell, img_h) if square else cell
    for i, (key, title, caption) in enumerate(items):
        cx = x + i * (cell + gap)
        image(slide, key, cx + (cell - img_w) / 2, y, img_w, img_h)
        text(
            slide,
            cx,
            y + img_h + 0.08,
            cell,
            0.5,
            [(title, 12, WHITE, True, 0), (caption, 10, GREY, False, 1)],
            align=PP_ALIGN.CENTER,
        )


# ---------- the slide ----------


def build_slide():
    CACHE.mkdir(exist_ok=True)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = NAVY
    rect(slide, MSO_SHAPE.RECTANGLE, 0, 0, 0.12, 7.5, RED)

    # header
    tf = slide.shapes.add_textbox(
        Inches(0.4), Inches(0.25), Inches(9), Inches(0.8)
    ).text_frame
    tf.margin_left = tf.margin_top = 0
    for s, colour in (("Hi, I'm ", WHITE), ("Mikołaj", YELLOW)):
        r = tf.paragraphs[0].add_run()
        r.text = s
        r.font.name, r.font.size, r.font.bold, r.font.color.rgb = (
            FONT,
            Pt(40),
            True,
            colour,
        )
    text(
        slide,
        0.4,
        1.0,
        12,
        0.35,
        [("22  ·  maths person  ·  AI student  ·  AI Engineer", 16, GREY, False, 0)],
    )

    # row 1
    y, h = 1.55, 2.75
    tile(slide, 0.4, y, 3.1, h, "ORIGIN")
    pairs = [
        ("FROM", "Strzelin"),
        ("LIVE IN", "Wrocław"),
        ("HIGH SCHOOL AT WUST", "Academic School · 2022"),
        ("STUDY", "AI · 5th year · WUST"),
    ]
    lines = []
    for label, value in pairs:
        lines += [
            (label, 9, GREY, True, 7 if lines else 0),
            (value, 15, WHITE, True, 0),
        ]
    text(slide, 0.65, y + 0.62, 2.7, 2.0, lines)

    tile(slide, 3.7, y, 4.9, h, "WORK")
    text(
        slide,
        3.95,
        y + 0.6,
        4.4,
        1.1,
        [
            ("AI Engineer", 22, WHITE, True, 0),
            ("working since August 2025", 12, YELLOW, False, 0),
            ("Accenture  →  Connells Group", 15, WHITE, False, 3),
            (
                "UK property services  ·  HQ: Milton Keynes  ·  remote",
                12,
                GREY,
                False,
                2,
            ),
        ],
    )
    image(slide, "accenture_logo", 3.95, y + 1.8, 1.9, 0.75)
    rect(slide, MSO_SHAPE.RIGHT_ARROW, 5.98, y + 2.0, 0.35, 0.35, YELLOW)
    image(slide, "connells_logo", 6.45, y + 1.8, 1.9, 0.75)

    tile(slide, 8.8, y, 4.133, h, "FORMULA 1")
    text(
        slide,
        9.05,
        y + 0.6,
        3.7,
        0.7,
        [
            ("Red Bull Racing", 20, WHITE, True, 0),
            ("fan since 2022  ·  HQ: Milton Keynes", 12, GREY, False, 2),
        ],
    )
    image(slide, "redbull_car", 9.05, y + 1.35, 3.633, 1.2)

    # row 2
    y, h = 4.5, 2.75
    tile(slide, 0.4, y, 6.9, h, "HOBBIES")
    text(
        slide,
        2.6,
        y + 0.22,
        4.45,
        0.3,
        [("+ computer games  ·  guitar", 12, GREY, False, 0)],
        align=PP_ALIGN.RIGHT,
    )
    gallery(
        slide,
        0.65,
        y + 0.62,
        6.4,
        [
            ("running", "Running", "Night Half Marathon · < 2 h"),
            ("hiking", "Hiking", "Tatras · Sudetes · Dolomites"),
            ("skiing", "Skiing", "Alta Badia, Italy"),
        ],
        img_h=1.45,
    )

    tile(slide, 7.5, y, 5.433, h, "MUSIC")
    gallery(
        slide,
        7.75,
        y + 0.62,
        4.933,
        [
            ("album_dire_straits", "Dire Straits", "London · Mum"),
            ("album_arctic_monkeys", "Arctic Monkeys", "Sheffield · me"),
            ("album_metallica", "Metallica", "USA, not UK! · Dad"),
        ],
        img_h=1.45,
        square=True,
    )

    prs.save(HERE / "about_me.pptx")


# ---------- the printable script ----------

SCRIPT = [
    (
        "1. Hello and where I'm from  (≈ 45 s · ORIGIN)",
        [
            "Hi everyone! For those who don't know me yet, I'm Mikołaj and I'm 22. I'm originally from Strzelin, "
            "a small town about 40 kilometres south of Wrocław, but these days I live in Wrocław itself.",
            "I've always been a maths person – numbers have simply made sense to me. In 2022 I finished the Academic "
            "Secondary School run by Wrocław University of Science and Technology, and then, quite naturally, I stayed "
            "at the same university. I'm now a fifth-year Artificial Intelligence student at the Faculty of Information "
            "and Communication Technology.",
        ],
    ),
    (
        "2. My job  (≈ 1 min · WORK)",
        [
            "Alongside my studies, I work. Since August 2025 I've been an AI Engineer at Accenture – together with "
            "Marcin, who's sitting right here.",
            "Sadly, I'm parting ways with Accenture this Tuesday, because I'm starting a new job at Connells Group. "
            "Connells is one of the leading property services providers in the UK – in other words, an estate agency, "
            "as the British would say. I'll still be an AI Engineer, but I'll be working fully remotely from Poland and "
            "flying over to the UK every few months – probably once a quarter.",
            "Their head office is in Milton Keynes, about 50 miles north-west of London. And what else is in Milton "
            "Keynes, you might ask? I'll come back to that at the end.",
        ],
    ),
    (
        "3. Hobbies  (≈ 1 min · HOBBIES)",
        [
            "When I'm not in front of a computer, I try to stay active. First of all, I run. This photo is from the "
            "Wrocław Night Half Marathon – I finished in just under two hours, and to be honest, I nearly coughed up "
            "a lung at the finish line.",
            "I also love hiking, mostly in the Polish mountains – the Tatras and the Sudetes. In winter I go skiing in "
            "Alta Badia, in the Italian Dolomites. This year I went back there without my skis and hiked the very "
            "slopes I usually ski down – it was a completely different experience.",
            "And when I just want to unwind, I play computer games or pick up my guitar – which brings me to music.",
        ],
    ),
    (
        "4. Music  (≈ 1.5 min · MUSIC)",
        [
            "My taste in music is a bit of a family affair: two British bands and one American one.",
            "Let's start with my mum. Her all-time favourite song is ‘Tunnel of Love’ by Dire Straits, a London band. I used to be "
            "able to play it, but these days I only remember the outro solo – which, to be fair, is the best part.",
            "Then there's me. My band is Arctic Monkeys – they're from Sheffield, and my favourite album is their debut, "
            "‘Whatever People Say I Am, That's What I'm Not’. I can play a few of their songs as well.",
            "And finally, my dad – and here's the exception. He's the reason I listen to Metallica, and Metallica are "
            "not British at all – they're an American band from California. Back in the 90s he was a proper fan – long "
            "hair, denim jacket, the whole lot. Nowadays he's more of a ‘Black Album’ listener, but I've picked up "
            "where he left off: I absolutely adore Metallica, and this year I saw them live in Chorzów. My favourite "
            "album is ‘Ride the Lightning’, and ‘Fade to Black’ is, for me, one of the most stunning songs they've ever "
            "written.",
        ],
    ),
    (
        "5. Formula 1  (≈ 45 s · FORMULA 1)",
        [
            "So, back to Milton Keynes. It's home to Red Bull Racing, my favourite Formula 1 team. I've been supporting "
            "them since 2022, so yes – I'm one of those ‘post-2021’ fans who got hooked after that dramatic season "
            "finale in Abu Dhabi. But that's fine – everyone has to start somewhere.",
            "And here's the best part: my new company's head office is in the very same town. So every time I fly over, "
            "I'll be just down the road from the Red Bull factory – maybe one day I'll finally get to visit it.",
        ],
    ),
    (
        "6. Wrap-up  (≈ 10 s)",
        [
            "So that's me in a nutshell: maths, AI, running, mountains, guitar and Formula 1. Thanks for listening – "
            "any questions?",
        ],
    ),
]

CORRECTIONS = [
    (
        "Always have been math-oriented person.",
        "I've always been a maths-oriented person.",
        "Subject needed; word order ‘have always been’; article ‘a’; British ‘maths’.",
    ),
    (
        "I graduated from acedemic high school of the WUST",
        "I finished the Academic Secondary School run by WUST",
        "Spelling ‘academic’. In the UK it's ‘secondary school’, and you ‘finish’ or ‘leave’ school – you "
        "‘graduate’ from university.",
    ),
    (
        "What do I like: running, …",
        "As for my hobbies, I'm into running, …",
        "Sounds more natural in speech.",
    ),
    (
        "playing computer games and guitar",
        "playing computer games and the guitar",
        "British English normally uses ‘the’ with instruments.",
    ),
    (
        "I am a big big fan of some british bands",
        "I'm a huge fan of a few British bands",
        "Nationalities take a capital letter; ‘huge’ instead of ‘big big’.",
    ),
    (
        "Tunnel of Love which I had known how to play sometime ago - only remember the outro solo",
        "‘Tunnel of Love’, which I used to be able to play – now I only remember the outro solo",
        "‘used to be able to’ for an ability you've lost; ‘had known’ is the wrong tense; subject needed before "
        "‘remember’.",
    ),
    (
        "I listen to the Arctic Monkeys band",
        "I listen to Arctic Monkeys",
        "No ‘the … band’ – the name is enough. British English treats bands as plural: ‘Arctic Monkeys are…’.",
    ),
    ("My favourtie album", "My favourite album", "Spelling."),
    (
        "I know some songs from them as well",
        "I can play a few of their songs as well",
        "‘their songs’ / ‘songs by them’, not ‘from them’; ‘can play’ is clearer than ‘know’.",
    ),
    (
        "my father's influence is Metallica - he's not a great fan, also the Black Album enjoyer",
        "Metallica came from my dad. He isn't a massive fan any more – more of a ‘Black Album’ listener",
        "‘enjoyer’ is internet slang; ‘any more’ shows the change since the 90s.",
    ),
    (
        "I do adore Metallica right now, was at the Concert in Chorzów this year",
        "These days I absolutely adore Metallica – I saw them live in Chorzów this year",
        "Subject needed before ‘was’; ‘concert’ is lower case; ‘saw them live’ is the natural phrase.",
    ),
    (
        "Ride the lightning … awesome song",
        "‘Ride the Lightning’ … a stunning / brilliant song",
        "Capitalise titles; ‘awesome’ sounds American – ‘brilliant’ is the British favourite.",
    ),
    (
        "I have been working since August 2025, with Marcin here my work colleague",
        "I've been working at Accenture since August 2025, alongside Marcin, who's sitting right here",
        "Name the company; ‘alongside’ plus a ‘who’ clause reads more smoothly.",
    ),
    (
        "Sadly I part ways with current employer - Accenture this tuesday as I start a new job at Connels Group",
        "Sadly, I'm parting ways with my current employer, Accenture, this Tuesday, as I'm starting a new job at "
        "Connells Group",
        "Present continuous for fixed plans; ‘my’ is needed; days of the week take a capital; the company is "
        "spelt ‘Connells’.",
    ),
    (
        "real estate agency in the UK",
        "estate agency",
        "‘Real estate’ is American; in the UK it's ‘estate agency’ / ‘estate agent’.",
    ),
    (
        "The headquarters are near Luton and quite close Milton Keynes",
        "Its head office is in Milton Keynes",
        "Missing preposition: ‘close to’. Fact check: Connells' website gives Witan Gate House, Milton Keynes as "
        "its head office.",
    ),
    (
        "What is in Milton Keynes you might ask",
        "And what's in Milton Keynes, you might ask?",
        "Comma before ‘you might ask’, and a question mark.",
    ),
    (
        "my beloved F1 team which i have been supporting since 2022",
        "my beloved F1 team, which I've been supporting since 2022",
        "Capital ‘I’; comma before a non-defining ‘which’.",
    ),
    (
        'yeah I am this "after 2021" guy. But thats okay.',
        "Yes, I'm one of those ‘post-2021’ fans. But that's fine.",
        "‘one of those …’ is the idiomatic pattern; apostrophe in ‘that's’.",
    ),
    (
        "I am on fifth year of Artificial Inteligence field",
        "I'm a fifth-year Artificial Intelligence student / I'm in my fifth year of AI",
        "‘in my … year’, not ‘on’; spelling ‘Intelligence’; ‘field’ isn't needed.",
    ),
    (
        "I will work fully remote … Quaterly perhaps",
        "I'll be working fully remotely … probably once a quarter",
        "Adverb ‘remotely’ after a verb; spelling ‘quarterly’.",
    ),
    (
        "i threw my lungs out",
        "I nearly coughed up a lung",
        "The English idiom for an exhausting effort (informal).",
    ),
    (
        "Hiked the slopes I ski on",
        "I hiked the very slopes I usually ski down",
        "Subject needed; ‘the very’ adds emphasis.",
    ),
]

PHRASES = [
    ("a maths person", "someone naturally good at, or fond of, maths"),
    ("alongside my studies", "at the same time as studying"),
    ("to part ways (with someone)", "to stop working together"),
    ("to fly over", "to travel by plane, usually to a nearby country"),
    ("once a quarter", "every three months"),
    ("to cough up a lung", "(informal) to be completely out of breath"),
    (
        "knackered",
        "(British, informal) exhausted – ‘After the half marathon I was absolutely knackered.’",
    ),
    ("to unwind", "to relax"),
    ("a family affair", "something the whole family is involved in"),
    ("to be fair", "(very British) used to add a balancing point"),
    ("a proper fan", "(British, informal) a real, serious fan"),
    ("the whole lot", "(British, informal) everything"),
    ("to pick up where someone left off", "to continue what someone else started"),
    ("to get hooked (on something)", "to become very keen on something"),
    ("just down the road", "very close by"),
    ("in a nutshell", "in short"),
]

BRITISH = [
    ("maths", "math"),
    ("mum", "mom"),
    ("secondary school", "high school"),
    ("estate agent / estate agency", "realtor / real estate agency"),
    ("favourite", "favorite"),
    ("brilliant", "awesome"),
    ("holiday", "vacation"),
    ("flat", "apartment"),
]


def _table(doc, header, rows):
    t = doc.add_table(rows=1, cols=len(header), style="Table Grid")
    for cell, h in zip(t.rows[0].cells, header):
        cell.paragraphs[0].add_run(h).bold = True
    for row in rows:
        for cell, value in zip(t.add_row().cells, row):
            cell.text = value


def build_script():
    doc = Document()
    for s in doc.sections:
        s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(2)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = DocPt(12)

    doc.add_heading("About me – talk script (about 5 minutes)", 0)
    doc.add_paragraph(
        "Read it aloud a few times, then speak from the keywords on the slide rather than reading."
    )
    for title, paragraphs in SCRIPT:
        doc.add_heading(title, 2)
        for p in paragraphs:
            doc.add_paragraph(p)

    doc.add_page_break()
    doc.add_heading("Corrections to your notes", 1)
    _table(doc, ["Your version", "Better", "Why"], CORRECTIONS)
    doc.add_heading("Useful phrases (C1, informal)", 1)
    for phrase, meaning in PHRASES:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(phrase).bold = True
        p.add_run(f" – {meaning}")
    doc.add_heading("British vs American", 1)
    _table(doc, ["British", "American"], BRITISH)
    doc.save(HERE / "talk_script.docx")


PAGE = """<!DOCTYPE html>
<html lang="en-GB">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Hi, I'm Mikołaj</title>
<style>
  html, body {{ margin: 0; height: 100%; background: #0b163a; }}
  body {{ display: flex; align-items: center; justify-content: center; }}
  img {{ max-width: 100vw; max-height: 100vh; }}
</style>
</head>
<body>
<img src="data:image/webp;base64,{data}" alt="Hi, I'm Mikołaj – origin, work, Formula 1, hobbies and music">
</body>
</html>
"""


def build_site():
    # Quick Look (macOS) renders the slide; the image is inlined so the site is a single file
    subprocess.run(["qlmanage", "-t", "-s", "2560", "-o", str(CACHE), str(HERE / "about_me.pptx")],
                   check=True, capture_output=True)
    buf = io.BytesIO()
    Image.open(CACHE / "about_me.pptx.png").convert("RGB").save(buf, "WEBP", quality=88)
    SITE.mkdir(exist_ok=True)
    (SITE / "index.html").write_text(PAGE.format(data=base64.b64encode(buf.getvalue()).decode()))


if __name__ == "__main__":
    build_slide()
    build_script()
    build_site()
    print("Saved about_me.pptx, talk_script.docx and docs/index.html")
