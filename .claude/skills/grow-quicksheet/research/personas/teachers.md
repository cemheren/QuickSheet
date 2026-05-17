# Persona 10 — Teachers / educators (K–12, college instructors, tutors)

Status: **done** · Last revised: 2026-05-16

## TL;DR (5 lines)

- ~3.8M US K–12 teachers + ~1.5M college faculty/adjuncts + uncountable independent tutors. **The most data-handling-skeptical of any persona** — burned by ed-tech SaaS sprawl, FERPA-paranoid, school-IT-locked-down.
- Single classroom laptop + projector. Wallpaper is *not* the right surface during teaching (projector shows slides) — but **the prep desktop and the between-classes desktop are perfect**.
- Highest-leverage wallpaper cells: this period's class strip, grading queue with progress bars, attendance summary, parent-contact log, lesson-plan-due countdowns, break timer, IEP/accommodation reminders.
- Where to seed: r/Teachers, r/TeachingResources, r/Professors, r/AskAcademia (faculty side), r/edutech, EduTwitter (still active for academics), Cult of Pedagogy network, ChronicleVitae.
- Direct competition: Google Classroom (forced by district), Canvas/Blackboard (LMS), Schoology, gradebook SaaS (Aeries, PowerSchool). **None on the wallpaper. None of them give the teacher a personal between-classes glance surface.**

## 1. Profile

- US: ~3.8M K–12 (NCES 2024) + ~1.5M college instructors + adjuncts (most underpaid, technically-engaged subset). Globally many millions more.
- Hardware: school-issued Chromebook or aging Dell for K–12; personal laptop + classroom PC + projector. College faculty: own MacBook + office desktop + classroom podium.
- Day-shape: 6–8 class periods (K–12) or 2–4 lecture slots (college) with prep/grading/parent-comm sandwiched between.
- Buying brain: "Does this respect student data? Will IT block it? Can I install it without admin rights?" *Local-first, CSV-first, offline-capable is exactly the right answer.*
- Cultural identity: undervalued, overworked, tool-fatigued. **High goodwill for tools made *for* them (not for the district admin layer).**
- Stress moment: Sunday-night prep + report-card season. The wallpaper grid earns its keep here.

## 2. Why their desktop is wasted

What lives there now:

- District-mandated wallpaper (school logo).
- Browser tabs: Google Classroom, Canvas, Aeries, district email, parent-comm portal, NYTimes lesson source.
- A grading spreadsheet downloaded from the LMS (CSV — speaks our language).
- Sticky notes labelled "call Jamie's mom Mon."
- Possibly a Word doc with the period schedule.

What QuickSheet replaces: the *intent* of opening the LMS to "see where I am with grading." With the wallpaper, **all grading queues, attendance, and parent-contact logs are painted on the desktop the moment the teacher closes a student's submission.** Between-class transitions become orientation moments, not lost minutes.

Critical for this persona: **data lives in a local CSV.** That alone passes the FERPA gut check most ed-tech tools fail. "Your gradebook is a file in your folder, not a SaaS in another country."

## 3. Glanceable data they actually want behind windows

1. **Today's period strip** — "P1 9:15 Algebra II", "P2 10:10 Pre-Calc", current period highlighted, time until next bell.
2. **Grading queue** — N rows: assignment name + class + count-ungraded + days-since-due. Red if overdue.
3. **Attendance summary** — today's absent count per class.
4. **Parent-contact log queue** — "Call Jamie's mom re: missing hw" with last-touched dates. Stale items dim.
5. **IEP / accommodation reminders** — flagged students for the next period.
6. **Lesson-plan deadlines** — next-week submissions to admin, observation prep.
7. **Bell timer** — countdown to next class transition.
8. **Snack-and-water reminders** — yes, seriously. Teachers forget to hydrate. Pomo-like cell.
9. **Personal-prof side** — tenure-clock papers-in-review count, grant deadlines, conference-CFP queue.

## 4. Candidate extensions / desktop-only features (ranked)

| Rank | Item                                       | Type | Cost | Hit prob | Why                                                                          |
|------|--------------------------------------------|------|------|----------|-------------------------------------------------------------------------------|
| 1    | `schedule:` ICS reader *(also students)*    | ext  | low  | high     | Same ext as students. Period strip + bell timer fall out for free.           |
| 2    | `grader:` local CSV gradebook summary       | ext  | low  | high     | Reads `gradebook.csv` (class, student, assignment, score); summarises queues. |
| 3    | **Value-driven cell colour**                | feat | low  | high     | Red overdue grading, missing hw. 10/10 personas. *Build first thing post-research.* |
| 4    | **Per-cell ticking timer**                  | feat | low  | high     | Bell timer + lesson-segment timer + snack reminder. Already 6/10.            |
| 5    | **In-cell progress bar prefix**             | feat | low  | high     | "23/30 graded" → `▓▓▓▓░ 23/30`. 5/10 personas.                              |
| 6    | `attend:` attendance CSV summariser          | ext  | low  | med-high | Reads `attendance.csv`; today/this-week absences per student.                |
| 7    | `iep:` accommodations highlighter           | ext  | low  | med      | Reads `students.csv` with accommodation flags; per-period reminder cell.    |
| 8    | `canvas:` *(also students)* LMS pull          | ext  | med  | high     | Same ext serves teacher side (gradebook pull). Hits 1.5M+ US college instructors. |
| 9    | `gclassroom:` Google Classroom assignments  | ext  | high | high     | Google OAuth heavy. Defer; pair with k12 specialisation later.               |
| 10   | `cfp:` academic CFP deadlines (WikiCFP RSS) | ext  | low  | med      | Faculty-side; reads WikiCFP RSS for a discipline.                            |

The 3 recurring features (value-colour, timer, progress-bar) **collectively serve every persona in the slate**. Teachers reinforce all three. After research-phase end these are the obvious Bucket E start.

## 5. Where they hang out

- **r/Teachers (1M), r/TeachingResources, r/AskAcademia, r/Professors (60k), r/AdjunctLife, r/k12sysadmin**.
- **r/edutech / r/edtech** — adjacent; receptive to local-first tools.
- **EduTwitter / EduBluesky** — still very active. Hashtags #edchat, #ntchat, #flipclass.
- **Cult of Pedagogy** (Jennifer Gonzalez, ~1M+ teachers reached monthly through podcast + blog).
- **The Chronicle of Higher Ed** + **ChronicleVitae** — faculty side; tool features land here occasionally.
- **Edutopia** — K–12 side.
- **Discords/Slacks:** TeachersConnect, Edutopia community, university-specific faculty Slacks.
- **People to be visible to:** Jennifer Gonzalez (Cult of Pedagogy), Larry Ferlazzo (EdWeek blogger), Vicki Davis (CoolCatTeacher). All publish tool roundups; local-first FERPA-friendly tools are catnip.

## 6. Discoverability hooks

Hero image = **a teacher prep desk**. Classroom PC with Google Slides open in browser. Wallpaper to the side: QuickSheet grid with today's period strip ("P3 12:15 starts in 8min"), grading queue ("Algebra Quiz: 23/30 graded"), attendance summary, "call Jamie's mom" in red.

Headlines that land:

- "I built a free, open-source desktop dashboard for teachers — your gradebook is just a CSV"
- "Local-first lesson planning + grading queue on your wallpaper (FERPA-friendly, no SaaS)"
- "Sunday-night prep got quieter: my grading queue and class schedule are on my wallpaper now"

Avoid: ed-tech bro language, "AI assistant for teachers" framing, anything claiming to replace the LMS. Frame as *complement* to whatever the district uses.

## 7. Implications (queue these)

1. **Build `grader:` extension** (Bucket F, post-research). Reads a local `gradebook.csv` — pure local, FERPA-safe by construction, immediate hit.
2. **Build `schedule:` ICS reader** (cross-persona: also students). Same ext, 2 personas.
3. **Bucket A: `examples/teacher-dashboard.csv` + `examples/professor-dashboard.csv` + `docs/for-teachers.md` + `docs/for-professors.md`.** Two starter sheets, K–12 and college. Heavy "data stays on your machine" framing.
4. **Bucket C draft: r/Teachers post + Cult of Pedagogy outreach pitch.** Save in `drafts/teachers-launch.md`. Lead with FERPA-friendly + local-first.

Cross-link:
- [writers-academics.md](writers-academics.md) — overlap on faculty side (papers-in-review, CFPs).
- [students.md](students.md) — shares `schedule:` extension.
- [lawyers.md](lawyers.md), [accountants.md](accountants.md) — same audit-trail / local-data-stays-here trust angle.

---

# Research phase summary (post-paper-10)

All 10 persona papers `done`. Highest-leverage findings:

**The trinity** — three features that recur across 8–10 personas. Build these first when research phase ends:

1. **Value-driven cell colour** (`c?:>X=red,>Y=yellow,*=green: <value>`) — **10/10 personas.** ~40 LOC.
2. **Per-cell ticking timer prefix** — **6/10 personas** (lawyers/accountants/artists/GMs/writers/teachers). ~60 LOC.
3. **In-cell progress bar prefix** (`p: N/M` → `▓▓▓░░░ N/M`) — **5/10 personas** (writers/accountants/artists/gamedev/students/teachers). ~30 LOC.

**Top viral-action ranking** (after the trinity ships):

1. r/unixporn student-rice post (persona 9).
2. r/unixporn DM-screen rice (persona 5).
3. r/selfhosted "Homepage.io alternative, not in a tab" (persona 8).
4. r/battlestations trader multi-monitor (persona 7).
5. Cult of Pedagogy / Cara.app outreach (personas 10 + 4) — slower but high-trust.

**Top first-extension-to-build ranking** (after trinity ships):

1. `health:` HTTP probe (homelab, writes-its-own-screenshot).
2. `leetcode:` + `gh:` user-streak (students, viral combo).
3. `roll:` dice + `init:` initiative (gamedev DM bundle).
4. `gha:` GitHub Actions (SRE).
5. `bill:` billing timer ext + matching cell-timer feature (lawyers + accountants).

**Audience-specific docs to ship as soon as the matching extensions land:**
`docs/for-homelab.md`, `for-students.md`, `for-dms.md`, `for-traders.md`, `for-sre.md`, `for-lawyers.md`, `for-accountants.md`, `for-writers.md`, `for-academics.md`, `for-artists.md`, `for-indiedevs.md`, `for-teachers.md`, `for-professors.md`. Plus matching `examples/*.csv` starter sheets.

Research phase ends with this paper. Next /grow-quicksheet run resumes normal Bucket A/B/C/E/F selection; first build pick: the trinity Bucket E features.
