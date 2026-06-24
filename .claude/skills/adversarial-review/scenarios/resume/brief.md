# Résumé candidate — fill out resume.docx

Fill the attached Harvard résumé template (`resume.docx`) for the candidate below.
The template uses heading styles, tab stops that right-align dates, and bullet
lists — **preserve all of that styling**; only swap the placeholder text. Delete the
bracketed `[Note: ...]` helper notes once you've handled them. Leave the `[drawing]`
element alone (it's part of the template).

## Candidate

- **Name:** Priya Raman
- **Contact:** 14 Plympton St, Cambridge, MA 02138 • praman@college.harvard.edu • (617) 555-0192

## Education

- **Harvard University**, Cambridge, MA — A.B. in Computer Science, GPA 3.8. Graduating May 2027.
- **Lincoln High School**, Portland, OR — National Merit Finalist. Graduated June 2023.

## Experience

- **Northwind Robotics**, San Francisco, CA — Software Engineering Intern, Jun 2025 – Aug 2025
  - Built an automated test harness that cut nightly regression time by 40%.
  - Shipped a telemetry dashboard used by 30+ engineers to triage field failures.
- **Harvard SEAS**, Cambridge, MA — Research Assistant, Sep 2024 – May 2025
  - Implemented a data pipeline processing 2M sensor records per day.

## Leadership & Activities

- **Harvard Robotics Club**, Cambridge, MA — Project Lead, 2024 – present
  - Led a 12-person team to a 2nd-place finish at the regional autonomous-rover competition.

## Polish the layout (paragraph spacing & indentation)

The template is cramped. Once the text is in, tidy the paragraph layout:

- Add **12pt space before** each of the three section headings (Education,
  Experience, Leadership & Activities) so the sections are visually separated.
- Add **6pt space after** each experience/leadership **entry** line (the
  org/title/date lines) so entries breathe instead of butting together.
- Give the contact line directly under the name **6pt space after** as well.

These are paragraph properties — set them in place; you don't need to retype the
text. `docx read` shows the current spacing/indent as `docx:p … space-after="…"`
hints, and `docx read --ast` shows the exact values.

## Fit it on one page (page margins)

The default 1-inch margins waste space. Set the **page margins to 0.5 inch** on all
four sides so the résumé fits on a single page. This is page setup (a section/page
property), not a paragraph property — set it once for the whole document. `docx read`
surfaces page geometry as a `docx:page … margins="…"` note.

## What "done" looks like

Every placeholder (name, contact, school, degree, dates, position, bullets) is
replaced with the candidate's real details, the `[Note: ...]` helper text is gone,
the section headings (Education / Experience / Leadership & Activities) keep their
heading style, the dates still right-align against the tab stop, the layout has
breathing room — space before the section headings and after the entries — and the
page margins are 0.5 inch so it fits on one page.
