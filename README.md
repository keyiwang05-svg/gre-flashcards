# GRE Flashcards

A browser-based GRE vocabulary trainer built with **React, TypeScript, and Vite**. This independent project combines a working study product with deliberately designed quiz logic, browser-based persistence, automated tests, and product analytics.

**[Try the live app](https://greflashcards.fun)** · **[Explore the code](https://github.com/keyiwang05-svg/gre-flashcards)**

## Why I built it

While preparing for the GRE, I wanted a practice tool that could support repeated vocabulary review without making synonym questions unfair. A conventional multiple-choice generator can accidentally treat another valid synonym as a wrong answer. Practice sessions also need to survive refreshes and changes in the selected word list.

The project grew from a personal study tool into a deployed application. I used actual usage patterns to improve entry into practice, while building and testing the underlying quiz logic.

## What the app does

- **Vocabulary flashcards:** study and review GRE words, with different study modes, filters, and progress tracking.
- **Synonym-pair practice:** generate multiple-choice questions from word pairs, with frequency-based practice pools and a wrong-answer review workflow.
- **Daily study workflow:** set goals, review due items, and revisit mistakes.
- **Import and export:** load vocabulary from JSON, CSV, or Excel files; extract pairs from PDFs that match the supported layout; export study data as JSON.
- **Browser-local progress:** persist study data and resume flashcard or pair-quiz sessions.
- **Product analytics:** instrument navigation and answer events to understand where users encounter friction.

## Engineering highlights

### 1. A synonym-aware quiz engine

The quiz engine is separated from the main UI in [`src/lib/pairQuiz.js`](src/lib/pairQuiz.js). The core design decisions are:

1. **Normalize and deduplicate pairs.** Unicode NFKC normalization, trimming, case folding, and order-independent pair keys identify duplicate pairs, including reversed entries.
2. **Avoid misleading distractors.** For each prompt, the generator excludes the correct answer, the prompt itself, and directly known synonyms drawn from the full pair library—not just the currently selected practice tier.
3. **Cover the selected pool before repeating it.** Each round uses a shuffled queue of unique pair keys. A completed round starts a new queue, avoiding an immediate repeat of the previous round's last item when possible.
4. **Resume safely.** The saved round contains JSON-serializable state. The engine validates queue membership and saved choices before restoring a question; changes to synonym relationships can invalidate old distractors.
5. **Keep the logic testable.** Randomness can be injected, allowing deterministic tests of coverage, ordering, and boundary conditions.

A deliberate limitation: distractor filtering uses **directly recorded synonym relationships**, not a transitive synonym closure. If A is related to B and B to C, the app does not automatically assume A and C have the same meaning.

### 2. Persistence and recovery

The application stores the main study state in **IndexedDB**, with a **localStorage fallback** if IndexedDB is unavailable or a write fails. Smaller UI and quiz-session snapshots are stored separately in localStorage.

This design allows users to return to ongoing practice without an account, but **progress is device/browser-local**: clearing browser data can remove it, and cross-device syncing is not implemented. The JSON export provides a manual backup of study data; the current import path accepts word/pair records rather than restoring every exported app-level statistic and setting.

See [`src/GREFlashcardsApp.tsx`](src/GREFlashcardsApp.tsx) for storage, import, and export handling.

### 3. Data ingestion

The app ships with built-in vocabulary and pair data in [`public/data/`](public/data/). It also supports:

| Input | Current handling |
| --- | --- |
| JSON | Word arrays, pair arrays, or objects containing `words` / `sixChoicePairs` |
| CSV | Vocabulary rows, including a supported Chinese multi-row vocabulary layout |
| XLS / XLSX | Spreadsheet rows parsed into vocabulary records |
| PDF | Text extraction and parsing for a specific structured synonym-pair layout |

Imported entries pass through normalization and merging/deduplication logic. **The PDF importer is format-specific; it is not a general PDF or OCR parser.**

### 4. Analytics-driven iteration

Event tracking is implemented in [`src/lib/analytics.ts`](src/lib/analytics.ts) using PostHog, alongside Vercel Analytics. Events include quiz starts, correct/incorrect answers, section navigation, and homepage practice clicks.

The PostHog configuration disables automatic capture and session recording. Event properties are filtered to exclude selected word-, answer-, and search-related fields; this describes the implemented instrumentation, not a comprehensive privacy audit.

**Product usage snapshot (November 2026, based on project analytics):**

- **1,000+ users** reached by the deployed learning tool.
- **26,000+ quiz-response events** recorded.
- **50% homepage-to-quiz click-through rate** observed after introducing a guided homepage.

The 50% figure is the observed share of homepage visitors who entered the quiz, **not** a measured 50% improvement or an A/B-test effect.

### 5. Automated testing

The repository includes **12 Node.js test cases** in [`tests/pairQuiz.test.mjs`](tests/pairQuiz.test.mjs), focused on the pair-quiz engine. They cover, among other cases:

- Full-round coverage of **1,014 built-in synonym pairs**, without repeats inside a round.
- Valid multiple-choice options and avoidance of known synonym distractors.
- Deduplication across case differences, spaces, reversed pairs, and IDs.
- Saved-session restoration, including the final question of a round.
- Changed filters, invalid saved queues, and new synonym relationships.
- Empty and single-pair pools.

These are unit-level tests for the quiz engine; they are **not** a claim of full UI, browser, or end-to-end test coverage.

## Technology and code map

**Frontend:** React 19, TypeScript, Vite, Tailwind CSS, Framer Motion  
**Data and files:** IndexedDB, localStorage, SheetJS (`xlsx`), PDF.js  
**Analytics:** PostHog, Vercel Analytics  
**Testing:** Node.js built-in test runner

```text
src/
  GREFlashcardsApp.tsx     Main study UI, state, persistence and file import
  lib/
    pairQuiz.js             Synonym-pair question generation and session logic
    analytics.ts            Event tracking and property filtering
  components/ui/            Reusable UI primitives
public/data/                Built-in word and synonym-pair data
tests/pairQuiz.test.mjs      Pair-quiz engine tests
```

Most application logic is currently concentrated in one React component; splitting it into smaller domain modules is a future maintainability improvement.

## Run locally

Use a Node.js version compatible with the repository's Vite version, then:

```bash
git clone https://github.com/keyiwang05-svg/gre-flashcards.git
cd gre-flashcards
npm ci
npm run dev
```

Open the local URL printed by Vite. The app can load its bundled study data without configuring PostHog.

**Optional analytics setup:** create a `.env.local` file containing:

```dotenv
VITE_POSTHOG_KEY=your_posthog_project_key
VITE_POSTHOG_HOST=your_posthog_host
```

Do not put private secrets in Vite's `VITE_` environment variables, since they are exposed to client-side code.

Useful commands:

```bash
npm test          # Run pair-quiz unit tests
npm run lint      # Run ESLint
npm run build     # TypeScript build + production bundle
npm run preview   # Preview the production build locally
```

## Scope and next steps

This is an independently developed learning product, created through iterative product design and AI-assisted coding. The focus has been on making practice usable, preserving study state, and testing failure-prone quiz behavior.

Current boundaries include browser-local progress without account sync, format-dependent PDF imports, partial JSON restore behavior, and tests concentrated on the pair-quiz engine. Priorities for further development include modularizing the main component, expanding automated integration tests, and improving data portability.

**Author:** Keyi Wang  
**License:** No repository license has been added yet; availability of code on GitHub does not automatically grant redistribution rights to the software or bundled vocabulary content.
