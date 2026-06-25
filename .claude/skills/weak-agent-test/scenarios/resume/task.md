# Polish a résumé and make the section headings consistent

We have a Harvard résumé template (`resume.docx`) that still has placeholder text in
it. Fill it in for the candidate below, then tidy up the formatting so it looks
professional when opened. The template already has the right layout — heading styles,
tab stops that right-align the dates, bullet lists — so preserve all of that; only swap
the placeholder text for the real content. Edit the file directly.

## Candidate

- **Name:** Priya Raman
- **Contact:** 14 Plympton St, Cambridge, MA 02138 · praman@college.harvard.edu · (617) 555-0192

### Education

- **Harvard University**, Cambridge, MA — A.B. in Computer Science, GPA 3.8. Graduating May 2027.
- **Lincoln High School**, Portland, OR — National Merit Finalist. Graduated June 2023.

### Experience

- **Northwind Robotics**, San Francisco, CA — Software Engineering Intern, Jun 2025 – Aug 2025
  - Built an automated test harness that cut nightly regression time by 40%.
  - Shipped a telemetry dashboard used by 30+ engineers to triage field failures.
- **Harvard SEAS**, Cambridge, MA — Research Assistant, Sep 2024 – May 2025
  - Implemented a data pipeline processing 2M sensor records per day.

### Leadership & Activities

- **Harvard Robotics Club**, Cambridge, MA — Project Lead, 2024 – present
  - Led a 12-person team to a 2nd-place finish at the regional autonomous-rover competition.

## Make the section headings consistent

The most important fix: the three section headings (Education, Experience, Leadership &
Activities) should all share the same heading style. Right now there's an
inconsistency — **Education** and **Leadership & Activities** use the Heading 1 style,
but **Experience** was left as plain bold centered text. Fix it in two steps:

1. Give the **Experience** heading the same **Heading 1** style the other two sections
   use — so all three share one style.
2. Restyle the **Heading 1 style itself** — set it to **navy (color 1F3864), 13pt,
   bold, with 12pt space above** — in **one change**. Because all three headings share
   the style, editing the style's definition updates all three at once; don't reformat
   each heading by hand. (The tool can inspect and change a style's definition directly
   — find the verb that edits styles.)

## The rest of the polish

- **Remove the `[Note: ...]` helper text** throughout once you've handled it.
- **Leave the `[drawing]` element alone** — it's part of the template.
- **Add a little space after each entry line** (the org / title / date lines in the
  Experience and Leadership sections, and the contact line directly under the name) so
  entries breathe instead of butting together — about 6pt after each. Set this in
  place; you don't need to retype the text to change its spacing.
- **Set the page margins to 0.5 inch** on all four sides so everything fits on one
  page. This is a page setup change for the whole document, not a per-paragraph change.
- **Keep the right-aligned date tab stops and the bullet lists** exactly as they are.

## What "done" looks like

Every placeholder (name, contact, school, degree, dates, position, bullets) is filled
with the candidate's real details. The `[Note: ...]` helper text is gone. All three
section headings — Education, Experience, Leadership & Activities — look identical:
navy, bold, 13pt, with a little space above, because they all share one heading style
that was restyled in a single change. The dates still right-align. Entries have
breathing room between them. Page margins are 0.5 inch. The `[drawing]` element is
still there. When you open the file, it reads as a clean, well-spaced one-page résumé.
