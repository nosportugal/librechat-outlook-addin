# AGENTS.md

Outlook Add-in that sends emails to a LibreChat API for AI summarization and reply drafting, built with vanilla JS + Webpack + Office.js.

## Landmines & Boundaries

✅ Always:

- Add both `en-GB` and `pt-PT` translations to `src/i18n.js` for every user-facing string — **PT is the primary locale**
- Use `DOMPurify.sanitize()` on any HTML rendered from external input (API responses, email bodies)
- Route all identity and token acquisition through `src/auth.js`; the add-in uses Entra ID SSO via Nested App Authentication and has no API-key fallback
- Use [Conventional Commits](https://www.conventionalcommits.org/) for all commit messages
- Run `npm run build` to verify changes compile

⚠️ Ask first:

- `manifest.xml` — defines Office Add-in registration, URLs, ribbon buttons, and permissions; changes affect all deployment targets
- `webpack.config.js` — entry points and env-var injection; wrong edits break all three bundles
- `Dockerfile` — production image; changes affect deployment

🚫 Never:

- Commit `.env` files, Entra credentials, API keys, tokens, or other secrets (`.env` is gitignored; use `.env.example` for documentation)
- Add an API-key input, API-key fallback, or instructions asking users to obtain or paste an API key; authentication is handled by Entra ID NAA
- Use `item.body.setAsync()` in compose mode — it resets scroll position on Outlook Windows; use `prependAsync` exclusively (see `src/commands/commands.js` line 47)
- Remove or alter the invisible Unicode markers (`\u200B\u200C\u200B`) or HTML comment markers (`<!--ai-reply-start-->`) in `src/commands/commands.js` — they identify AI-generated content for replacement

## Commands

```bash
npm install          # install dependencies
npm run build        # production build → dist/
npm start            # dev server on https://localhost:3000 (HTTPS required by Office.js)
npm run dev          # alias for npm start
# npm run lint       # TODO: not yet available — eslint config not added yet
```

No test framework is configured yet.

## Testing

No tests exist.
When tests are added, document the framework and commands here.

## Code Style

Authentication boundaries:

- `src/auth.js` is the single owner of MSAL/NAA configuration, silent token acquisition, interactive fallback, and signed-in account lookup.
- `src/taskpane/` may display authentication state and start interactive sign-in when the host supports NAA.
- `src/commands/` must use silent token acquisition only because ribbon commands have no UI surface for a sign-in prompt.

Three entry points — keep them separate:

```
src/taskpane/  → main UI panel (summarize + settings)
src/commands/  → ribbon button handlers (reply generation, compose mode)
src/dialogs/   → prompt customization dialog
```

i18n pattern — every user-facing string uses a key in `src/i18n.js`:

```js
// ✅ correct — use translation keys
showError(t("notify.authFailed"));

// ❌ wrong — hardcoded user-facing string
showError("Sign-in failed. Please try again.");
```

HTML sanitization pattern:

```js
// ✅ correct — sanitize before rendering
const clean = DOMPurify.sanitize(html);

// ❌ wrong — raw insertion
element.innerHTML = apiResponse;
```

## Documentation Standards

All documentation lives in `docs/` (create if needed).
ATX headers, one sentence per line, relative links.
