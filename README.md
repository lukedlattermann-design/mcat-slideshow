# mcat-slideshow

Part of the **MCAT AI Tutoring System** — a three-skill suite for Claude Code that gives pre-medical students a complete, expert-validated MCAT preparation pipeline.

## What this skill does

Generates a complete, tutor-quality MCAT instructional slideshow as a `.pptx` file on any requested topic. Each deck follows a consistent structure:

- Title slide and session agenda
- Topic-block concept slides with key terms and mnemonics
- Worked examples for calculation-heavy topics
- Practice questions interspersed immediately after the content they test, numbered sequentially
- A clinical or physiological application slide
- A high-yield MCAT takeaways slide
- A final answer key with correct answers and one-line rationales for every question

Practice questions are sourced from verified AAMC repositories where possible. When primary sources are unavailable, original questions are written in AAMC format and clearly labeled as such.

Slideshows are saved to an `MCAT Presentations/` folder in your current working directory.

## Prerequisites

```bash
pip install python-pptx
```

## Install this skill

```bash
npx skills add lukedlattermann-design/mcat-slideshow
```

## Install the full three-skill system

For the complete workflow — schedule generation, daily check-ins, and on-demand slideshows — install all three skills:

```bash
npx skills add lukedlattermann-design/mcat-scheduler
npx skills add lukedlattermann-design/mcat-daily-checkin
npx skills add lukedlattermann-design/mcat-slideshow
```

## How the full system works

1. **`/mcat-scheduler`** — Run once at the start of your prep. Collects your details and generates a personalized week-by-week schedule as an `.xlsx` file.
2. **`/mcat-daily-checkin`** — Run every study day. Reads your schedule, tells you exactly what to study today, and asks if you are struggling with anything.
3. **`/mcat-slideshow [topic]`** — Triggered automatically by the daily check-in when you flag a struggling topic, or run manually at any time. Generates a tutor-quality `.pptx` slideshow with concept slides, worked examples, AAMC-style practice questions, and an answer key.

## Designed for

- Pre-medical students preparing for the MCAT who do not have access to professional tutoring
- MCAT tutors who want to automate schedule generation and session material preparation
