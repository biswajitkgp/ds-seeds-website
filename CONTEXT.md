# Project Context — D S Seeds Website

This file holds the stable facts about the project. Read it together with `HANDOVER.md` at the start of every new session (on any computer or Claude account).

## Client
- **Legal name:** D S Seeds (OPC) Private Limited
- **Industry:** Agricultural seeds
- **Relationship:** The owner approached Biswajit (Embetronix Technologies) to build the company website end to end.

## Service provider
- **Developer / consultant:** Biswajit Ghosh, Embetronix Technologies (Kharagpur, West Bengal, India)
- **Branding on client-facing documents:** Embetronix Technologies

## Brief (as given by the client)
- Dynamic website showing the product portfolio, company details and newsletters.
- **No e-commerce.** Informational / corporate website, but product **prices are displayed**.
- Main objective: make it easy for prospective buyers to discover the company and contact them.
- Must be mobile-optimised and fast-loading.
- Embetronix handles everything end to end: domain, hosting, design, build.
- The client is non-technical; client-facing documents avoid technical jargon (domain, hosting, etc.).

## Working agreements with Biswajit
- Ask clarifying questions before generating code or commands; never assume. Proceed only on explicit confirmation.
- Responses: structured, professor-like tone, forward-thinking; minimal emojis and bullet points.
- Procedural instructions one step at a time.
- Sync to GitHub (private repo) after every finished deliverable, or after a significant part of a big deliverable.
- Keep `HANDOVER.md` updated at each sync so work can continue from another computer or another Claude account.

## Repository layout
- `docs/01-requirements/` — requirement-gathering documents (questionnaire, client responses).
- `scripts/` — generators for documents (e.g. `build_questionnaire.js`, run with Node and the `docx` npm package).
- `HANDOVER.md` — current status, decisions, open questions, next steps.
- `CONTEXT.md` — this file.
