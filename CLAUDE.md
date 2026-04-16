# CLAUDE.md — Personal Preferences

These rules apply to all projects unless a project-specific CLAUDE.md overrides them.

---

## 1. Communication

- **Deviations**: If you need to deviate from stated requirements for any reason (technical limitation, ambiguity, better approach), **STOP and inform me in highlighted text before proceeding.** Do not silently deviate.
- **Ambiguity**: If any requirement is unclear or missing detail, ask a clarifying question before implementing. Do not guess.
- **Step gates**: After completing each meaningful step, stop and report what was done. Wait for explicit approval before starting the next step.
- **Commit gate**: After each step is coded and approved, commit all changes before moving on.

---

## 2. Code Quality

- TypeScript-first. No `.js` files in TS projects.
- No `any` types unless absolutely necessary — document with a comment explaining why.
- Use interfaces/types for all API shapes, DB models, and component props.
- Keep components small and focused. One component per file.
- Use custom hooks for shared logic (API calls, streaming, state, etc.).
- Run `npm run lint` (or equivalent) before considering any task done.
- Naming: concise in dense logic, descriptive for props, state, and functions.

---

## 3. UI/UX

- Match mockup screenshots exactly when provided. The visual reference is the final authority.
- Generous whitespace. Do not crowd elements.
- Border radius: 6–8px on cards, inputs, and badges.
- Minimal shadows. Prefer borders over shadows.
- Every screen must handle all four states: **loading**, **error**, **empty**, and **populated**.
- All interactive elements need **hover** and **focus** states.
- Streaming responses: show a typing/loading indicator before tokens arrive, then render tokens in real-time.

---

## 4. File Management

- Keep `README.md` updated with setup instructions and architecture overview.
- `.env.example` must list all required env vars with placeholder values (never real secrets).
- Gitignore: `node_modules/`, build output, `.env`, any local DB files, uploaded user content.

---

## 5. General Constraints

- Do not add features, screens, or UI elements not described in the requirements.
- Do not refactor surrounding code when fixing a bug or adding a targeted feature.
- Do not add comments or docstrings to code you did not change.
- Do not create helpers or abstractions for one-time use.
