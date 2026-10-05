---
name: iitp-slides
description: >-
  Builds and edits IITP'26 Computational Color Photography lecture decks by
  cloning Lec_1,2 Simple Light (python-pptx), not a new theme. Use when working
  on weeks/*/slides.md, weeks/*/build_lec*.py, Презентации/*.pptx, lecture
  slides, PPTX layout, or when the user asks to go through slides.
---

# IITP'26 lecture slides

Full spec: [templates/стиль-слайдов.md](../../../templates/стиль-слайдов.md).
**Reference deck:** `Презентации/Lec_1,2_intro2color.pptx` (taught L1+L2).
Do not clone `Lec_2_linear_model.pptx`. `Lec_1_intro2color.pptx` is gone.
Current builder: `weeks/03/build_lec3.py`.

Personal `pptx` (anthropics): unpack / thumbnail / schema QA only. Never pptxgenjs from scratch.

## Hard rules

- Clone `Lec_1,2_intro2color.pptx`, delete slides, keep masters and theme.
- Size **10″ × 5.625″**. Never 13.3″, Calibri, navy bar, or `Presentation()`.
- Course title IITP'26 / ИППИ. Do not copy TMSh / AIRI / MIPT / ITMO logos from the reference title.
- Fill the **body placeholder** (idx 1) on `TITLE_ONLY`. Do not make every slide out of `add_textbox`.
- Body **18 pt** Lato. Captions **11 pt** gray. Questions **30 pt** centered.
- Prefer figures already in `Lec_1,2` (BRDF diagram, LMS, Planck, CMF, locus, lettuce) over homemade matplotlib boxes.
- Do not commit `books/*.pdf`. Do not invent a new theme.

## Layout trap

`prs.slide_layouts` sees only the **first** master. That `TITLE_ONLY` has **no** header/footer rules.

Live slides use the **third** master’s `TITLE_ONLY` (`slideLayout23`): title + **body** + slide number + two `cxnSp` lines.

```python
P_CXNSP = "{http://schemas.openxmlformats.org/presentationml/2006/main}cxnSp"

def layout_named(prs, name):
    found = []
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name == name:
                found.append(layout)
    if name == "TITLE_ONLY":
        ruled = [ly for ly in found if ly._element.findall(f".//{P_CXNSP}")]
        if ruled:
            return ruled[0]
    return found[0]
```

After a rebuild, slide rels must point at `slideLayout23.xml`, not `slideLayout5.xml`.

Do not use `TITLE_AND_BODY` or `SECTION_HEADER`. Sections = ruled `TITLE_ONLY`, empty title, 30 pt in the body placeholder.

## Geometry (inches, live L1+L2)

| Box | left | top | width | height |
|---|---|---|---|---|
| Title | 0.355 | 0.463 | 9.183 | 0.695 |
| Body | 0.912 | 1.236 | 8.112 | 3.549 |
| Body if pic right | 0.912 | 1.236 | 4.677 | 3.549 |
| Picture right | 5.273 | 1.269 | 4.375 | 3.306 |
| Question | 0.912 | 2.153 | 8.112 | 1.481 |
| Caption | 0.447 | 5.188 | 9.1 | 0.387 |

Custom textboxes default to Calibri unless `font.name = "Lato"`. Do not assign `text_frame.text`.

## Filling (how the taught deck actually looks)

- One idea; a **figure holds the slide**.
- Typical: short body text + one picture on the right.
- Or one wide figure under the title + 11 pt source.
- Or two figures (locus, Planck), not a strip of three equal photos.
- Question slides between blocks.

## Workflow

1. Edit `weeks/NN/slides.md`.
2. Change `weeks/NN/build_lecN.py`; figures in `weeks/NN/figures/`.
3. Run from `C:\Users\Egor` (apostrophe in `IITP'26` breaks some PowerShell wrappers).
4. Verify `slideLayout23` and title box 0.355 / 0.463.
5. Walk slides with the user one at a time.

## QA

- [ ] Cloned `Lec_1,2_intro2color.pptx`
- [ ] 10 × 5.625
- [ ] Content → ruled `TITLE_ONLY` (`slideLayout23`)
- [ ] Body placeholder used for main text
- [ ] Lato, not Calibri
- [ ] Slide numbers intact
- [ ] No TMSh logos on the IITP title
- [ ] No program-of-the-course tables

## Files

| File | Role |
|---|---|
| `weeks/NN/slides.md` | Pair script |
| `weeks/NN/brief.md` | Pre-class mail |
| `weeks/NN/build_lecN.py` | Clone + fill |
| `Презентации/Lec_N_….pptx` | What is shown |
