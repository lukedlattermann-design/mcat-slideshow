---
name: mcat-slideshow
description: >
  Generate a tutor-quality MCAT slideshow presentation on a specific topic. Use when:
  (1) a student says they are struggling with a topic and wants a slideshow,
  (2) the user explicitly asks for a "tutor-quality slideshow" on a topic,
  (3) the mcat-daily-checkin skill triggers this after a struggle check.
  Output is a .pptx file saved to an `MCAT Presentations/` folder in the current working directory.
---

# MCAT Tutor-Quality Slideshow Generator

## Purpose

Generate a complete, ready-to-use MCAT tutoring slideshow on the requested topic. The output must
match the quality of the sample presentations in `mcat.sample.presentations/`. It will be used
directly by a professional MCAT tutor with students.

---

## Step 1 — Identify the Topic and Yield

Determine the topic from the invocation argument or from conversation context.

Look up the topic in the Kaplan chapter reference at
`references/kaplan-chapters.md` (included in this skill's directory). Find the chapter
name and star rating (✮ to ✮✮✮✮). Note the subject and yield.

If the topic spans multiple chapters, identify all relevant chapters.

---

## Step 2 — Scrape Topic Structure from Jack Westin

Search the web for: `site:jackwestin.com [topic name] MCAT`

Fetch the Jack Westin page for this topic. Extract:
- The organized list of subtopics / learning objectives for this MCAT topic
- Any links to related AAMC questions (note their IDs/references — you will find the actual text elsewhere)

Jack Westin URL pattern: `https://jackwestin.com/resources/mcat-content/[subject]/[topic]`

If Jack Westin has a page for this topic, use its subtopic structure as the backbone of the slide outline.

---

## Step 3 — Find AAMC-Style Practice Questions

### Priority order (always try in this order):

1. **AAMC questions** — search the web for exact AAMC question text:
   - Query: `"AAMC" "[topic]" MCAT practice question site:khanacademy.org`
   - Query: `AAMC MCAT "[topic]" practice question "[key term from topic]"`
   - Fetch Khan Academy MCAT pages for this topic: `https://www.khanacademy.org/test-prep/mcat`
   - Search r/MCAT for resource mentions: `site:reddit.com/r/mcat "[topic]" question resource`
     — follow any linked sources mentioned in comments (do NOT copy text from Reddit directly)

2. **High-quality third-party questions** — if AAMC questions cannot be found:
   - Search Khan Academy for passage-based questions on this topic
   - Look for questions that match AAMC style: 4 answer choices, science passage + discrete questions mix

### Collect 4–6 practice questions total:
- At least 1 passage-based set (one passage, 3–4 questions) if this is a science topic
- Include discrete (standalone) questions for high-yield facts
- For Psych/Soc topics: concept-application questions rather than calculation questions
- Always note the source of each question

---

## Step 4 — Build the Slide Outline

Practice questions are **interspersed throughout the deck** — they follow immediately after the topic
they test. Number all practice questions sequentially (Q1, Q2, Q3, ...) across the entire
presentation. The final slide is always an Answer Key.

**Overall structure:**

1. **Title slide** — topic name, subject area, "MCAT Tutoring Session" subtitle, yield indicator

2. **Agenda slide** — bullet list of all major topic blocks covered + estimated session time

3. **Topic block** (repeat for each major subtopic from the Jack Westin outline):

   a. **Concept slide(s)** — one per subtopic:
      - Clear title
      - Concise explanation (2–4 sentences)
      - Key terms bolded or in a definition box
      - Chemical equations / formulas / diagrams described (or placeholder label)
      - Mnemonic if one exists or can be created

   b. **Worked example slide** (only for calculation-heavy subtopics in Gen Chem, Physics, Biochem):
      - Title bar: "Example: [brief description]"
      - Step-by-step numbered list in body

   c. **Concept check slide** (optional — one per major topic area, not every subtopic):
      - Title bar: "Concept Check"
      - Single question prompt, Pt(20)
      - "Answer aloud before advancing" in italics Pt(12) at bottom
      - No answer shown on the slide — student answers aloud

   d. **Practice question slide(s)** — 1–2 questions that test THIS specific topic:
      - Label each: "Q[N]:" continuing the sequential numbering across the full deck
      - Discrete: question stem Pt(16) + answer choices A–D each on its own line Pt(14)
      - Passage-based: one slide for the passage text, then one slide per question labeled "Q[N]:"
      - Source noted at bottom Pt(10) italics: "Source: AAMC [ID]", "Source: Khan Academy",
        or "Source: AAMC-style (original)"

4. **Clinical / physiological application slide** — after all topic blocks; always present, even for
   Psych/Soc (connect to clinical psychology or health disparities)

5. **High-Yield MCAT Takeaways slide** — numbered list of top 8–10 facts/rules
   - Circle numbers ①②③... Pt(15), one item per line

6. **Answer Key slide** — ALWAYS the very last slide
   - Title bar: "Answer Key"
   - List every numbered practice question with its correct answer letter and a one-line rationale:
     e.g., "Q1: B — sp³ hybridization produces 4 electron groups, giving tetrahedral geometry"
   - Pt(14), one question per line
   - This is the only place in the deck where correct answers are revealed

**Question distribution:**
- Aim for 1–2 practice questions per topic block
- Total 4–8 practice questions across the deck
- At least 1 passage-based set if this is a science topic (passage slide + question slides)
- Questions must specifically test the concept immediately preceding them — not general topic trivia

---

## Step 5 — Generate the .pptx File

Use python-pptx to build the presentation. Run this via Bash.

### Setup

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# Slide dimensions: 10" x 5.625" (16:9 widescreen — matches sample presentations)
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)
```

### Subject color mapping (use for title bars and accents)

```python
SUBJECT_COLORS = {
    "General Chemistry":           RGBColor(0xEA, 0x99, 0x99),
    "Organic Chemistry":           RGBColor(0xFF, 0xE5, 0x99),
    "Physics":                     RGBColor(0x9F, 0xC5, 0xE8),
    "Biology":                     RGBColor(0xB6, 0xD7, 0xA8),
    "Biochemistry":                RGBColor(0xA4, 0xC2, 0xF4),
    "Behavioral Sciences":         RGBColor(0xD5, 0xA6, 0xBD),
    "Psychology":                  RGBColor(0xD5, 0xA6, 0xBD),
    "Sociology":                   RGBColor(0xD5, 0xA6, 0xBD),
}
```

### Helper: add a slide with title bar

```python
def add_slide(prs, layout_idx=6):
    layout = prs.slide_layouts[layout_idx]  # blank layout
    return prs.slides.add_slide(layout)

def add_title_bar(slide, title_text, subject, width=Inches(10), height=Inches(0.75)):
    from pptx.util import Inches, Pt
    txBox = slide.shapes.add_textbox(Inches(0), Inches(0), width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    # Fill the title bar shape with subject color (darker shade)
    from pptx.dml.color import RGBColor as RGB
    from pptx.oxml.ns import qn
    from lxml import etree
    fill = txBox.fill
    fill.solid()
    # Darken the subject color for contrast
    base = SUBJECT_COLORS.get(subject, RGB(0x67, 0x67, 0x67))
    fill.fore_color.rgb = RGB(
        max(0, base[0] - 40),
        max(0, base[1] - 40),
        max(0, base[2] - 40)
    )
```

### Slide construction rules

**Title slide:**
- Full-width colored header bar with topic name (white text, Pt(32), bold)
- Subtitle: "MCAT [Subject] — Tutoring Session" in Pt(18)
- Yield indicator: "Yield: ✮✮✮✮ (High)" in Pt(14) below subtitle

**Concept slides:**
- Title bar (colored, subject-matched) at top
- Body text box from y=0.85" to y=5.2" at x=0.3"
- Title text: Pt(22), bold, white on colored bar
- Body: Pt(16) regular, black, left-aligned
- Key terms: Pt(16), bold
- Equations: Pt(18), monospace-style (use Courier New or Calibri)
- Leave 0.4" margin on left and right

**Worked example slides:**
- Title bar: "Example: [description]"
- Step-by-step numbered list in body
- Each step on its own paragraph, numbered (1., 2., 3., ...)

**Concept check slides:**
- Title bar: "Concept Check" or "Pause and Practice"
- Single question prompt in large text (Pt(20))
- Add a text note at the bottom: "Answer aloud before advancing" in italics Pt(12)

**Practice question slides:**
- For passage: full passage text in Pt(13), then "Question X of Y:" header
- For discrete: question stem in Pt(16), then answer choices A–D each on their own line in Pt(14)
- Source noted at bottom in Pt(10) italics: "Source: AAMC [ID]" or "Source: Khan Academy"

**High-Yield Takeaways slide:**
- Title bar: "High-Yield MCAT Takeaways"
- Numbered list, Pt(15), each item one line
- Circle numbers ①②③... for visual consistency with sample presentations

### Save

```python
import os
output_dir = os.path.join(os.getcwd(), "MCAT Presentations")
os.makedirs(output_dir, exist_ok=True)
topic_slug = topic_name.replace(" ", "_").replace("/", "-")
output_path = f"{output_dir}/{topic_slug}.pptx"
prs.save(output_path)
print(f"Saved: {output_path}")
```

---

## Step 6 — Quality Check

After generating, verify:
- [ ] At least 15 slides (for a 1-hour session topic); at least 25 for a 1.5-hour topic
- [ ] Practice questions follow immediately after the topic they test (not grouped at the end)
- [ ] All practice questions are numbered sequentially Q1, Q2, Q3, ...
- [ ] Answer Key slide is the very last slide and lists all Q numbers with correct answers + rationale
- [ ] Takeaways slide is second-to-last (immediately before Answer Key)
- [ ] Practice questions total at least 4 across the deck
- [ ] At least one clinical/physiological application slide
- [ ] Topic title matches the Kaplan chapter name exactly
- [ ] Source attributed on every practice question slide

If any check fails, add the missing content before saving.

---

## Step 7 — Deliver

State the file path to the user:

> "Your slideshow on [topic] has been saved to: ./MCAT Presentations/[filename].pptx"
> "It contains [N] slides covering [list of main subtopics]. The practice questions are sourced from [AAMC/Khan Academy]."

---

## Notes

- **Never fabricate AAMC question text.** If you cannot find real AAMC questions, use Khan Academy
  questions or clearly labeled "AAMC-style" questions you write yourself that match the difficulty
  and format. Always label the source.
- **Psych/Soc topics**: If the student is weak in Psych/Soc, always add a note at the end:
  "Supplement this deck with the JackSparrow 2048 Anki deck for memorization."
- **High-yield first**: If the topic has ✮✮✮✮ rating, the takeaways slide should be especially
  thorough — these are the facts most likely to appear on the actual MCAT.
- **Session duration guidance**: include on agenda slide — roughly 15 slides ≈ 1 hour of tutoring.
