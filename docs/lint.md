# PV lint engine — design

Status: planned. This is the spec the implementation tracks against.

## Goal

Promote the deterministic ("mechanical") editorial checks — UK→US spelling, dash
normalization, "vs"→"versus", figure/table reference completeness, heading jumps, and so on —
from ad-hoc grepping into a first-class, repeatable PV capability, modelled loosely on the
*capabilities* of the Language Server Protocol (diagnostics, code actions, rename, document
symbols) without adopting the LSP wire protocol (which fights the Google Docs model — see
*Constraints*).

## The boundary: engine in PV, rules in context

PV must stay book-agnostic. Therefore:

- **PV (this repo) ships:** the lint engine, a small fixed set of rule *types* (primitives),
  and the result contract. It knows nothing about em-dashes, "buy versus build", or `Figure
  X-Y`.
- **The work's context repo ships the rules:** `context/<work>/lint-rules.toml` holds every
  actual rule — pattern, message, severity, scope, fix, and per-rule ambiguity policy. It is
  the executable companion to that work's `style_guide.md` (prose stays the human source of
  truth; the TOML is the machine-runnable subset).

PV is *passed* the rules file: `pv lint <doc-url> --rules <file>`. A user's specific needs
never enter `pv.py`.

## The result contract (three-valued) — foundational

Every mutating operation returns one of three statuses. This is introduced first, before
lint, because everything else inherits it.

- **`ok`** — done; payload describes what changed.
- **`error`** — genuine failure (auth, doc not found, malformed input); not resolvable by
  choosing.
- **`ambiguous`** — PV *could* act but there is a real choice it must not make unilaterally.
  It returns a well-formed decision request rather than a guess or a bare failure.

### `ambiguous` payload

```json
{
  "status": "ambiguous",
  "reason": "multiple_matches",
  "message": "human-readable explanation",
  "question": "Which occurrence did you mean?",
  "options": [
    {"id": 1, "paragraph_index": 30, "context": "…surrounding text…"},
    {"id": 2, "paragraph_index": 88, "context": "…surrounding text…"}
  ],
  "resolution": {"how": "re_call_with", "field": "occurrence",
                 "example": "pv edit <doc> <old> <new> --occurrence 2"}
}
```

`resolution` gives the caller the machine-actionable path back to success — this is a
protocol for getting unstuck, not a prettier error.

### The caller's two paths

On `ambiguous` the coding agent either:

- **self-resolves** — it has the context, so it re-calls with the disambiguating field
  (`--occurrence N`, a longer anchor, …); or
- **surfaces to the user** — it renders `question` + `options` into a disambiguation prompt.

The payload carries enough context to serve both.

### Where ambiguity legitimately arises

- `edit`/`link`/`style` **multiple matches** → `multiple_matches` (was a hard error).
- **no match** (text edited live, smart-quote / whitespace drift) → `no_match`, with nearest
  fuzzy candidates so the caller can recover.
- lint `--fix` whose range would **span a hyperlink or table boundary** → `unsafe_fix`
  (would otherwise destroy the link).
- `rename-term` occurrences in a different sense / inside a quote → `ambiguous_occurrences`.
- structural checks where non-sequential numbering *might* be intentional → flagged as a
  question, not an assertion.

No backwards-compatibility constraint (single-user tool): `ambiguous` is the default; the old
hard-error behaviour on `edit`/`link`/`style` is replaced outright.

## Rule file (`lint-rules.toml`, in context)

TOML — stdlib `tomllib`, no new dependency, supports comments and readable regex.

```toml
[[rule]]
id       = "uk-spelling-artefact"
type     = "replace"        # primitive PV implements
severity = "mechanical"     # mechanical | judgment | discussion
scope    = "prose"          # prose | code | all
pattern  = '\bartefact\b'
replace  = "artifact"
message  = "UK spelling; US house style is 'artifact'."

[[rule]]
id       = "tic-the-fact-that"
type     = "flag"           # diagnostic only, no safe autofix
severity = "judgment"
scope    = "prose"
pattern  = '(?i)the fact that'
message  = "Throat-clearing; rephrase to cut it."
on_ambiguity = "flag"       # flag (default) | first | error
```

### Rule types (the only book-agnostic part)

- `replace` — regex → replacement; autofixable.
- `flag` — regex → message; no autofix (judgment tier).
- `figure` / `table` / `listing` — structural: each numbered item needs a reference before
  it, a caption matching a configured format, and a sequential number. The format/ref regexes
  live in the rule file; the traversal lives in PV.
- `heading-jump` — a heading directly under a higher heading with no body between.
- `bullet-terminal-punct` — list items ending in a full stop.

`scope` is load-bearing: PV tags each paragraph prose/code (it already exposes `is_code`) so
`scope="prose"` rules skip code/monospace paragraphs — this is what keeps a blanket
double-space fix from mangling embedded JSON/code.

## Command surface

- `pv lint <doc-url> --rules <file>` → diagnostics JSON:
  `[{rule, severity, message, anchor_text, paragraph_index, suggested_fix?, autofixable}]`.
- `--fix` → apply autofixable `replace` rules in one batch (reverse-order or re-fetch between
  to survive index shifts; reuse the edit pipeline; any `unsafe_fix` returns `ambiguous`).
- `--comment` → post each diagnostic as an anchored sidebar comment (the `publishDiagnostics`
  analogue; PV already does anchored comments).
- `--severity mechanical` / `--rule <id>` filters.

Sibling capabilities (later): `pv rename-term`, `pv outline` (documentSymbol),
`pv refs <figure|term>`.

## Constraints (why full LSP is the wrong target on Google Docs)

- **No lines; shifting indices.** Docs addresses content by absolute offsets that move on
  every edit; LSP assumes stable `(line, char)`. PV already re-anchors on text, not indices —
  keep that.
- **Not plain text.** Tables, styles, links, comments, and tracked-change suggestions have no
  LSP text-document representation. `_extract_text` is a *rendered, lossy* projection
  (`[text](url)`, `" | "`-joined cells); a `replace` fix must never span a link/table
  boundary.
- **Network API, rate-limited.** Batch only; no diagnostics-on-keystroke.
- **Google Docs is the editor, edited concurrently.** The buffer changes underneath you;
  re-fetch-and-re-anchor is mandatory.
- **Nice fit:** diagnostics → comments, code-actions → edits. Both halves already exist.

## Phasing

1. **Result contract** — `ok`/`error`/`ambiguous` for `edit`/`link`/`style`, with
   `--occurrence` and fuzzy no-match candidates. (First commit; everything inherits it.)
2. **Lint engine** — `replace` + `flag` types + scope + `pv lint`/`--fix`, reading
   `lint-rules.toml`.
3. **Structural types** — `figure`/`table`/`listing`, `heading-jump`,
   `bullet-terminal-punct`; `--comment`.
4. **Sibling capabilities** — `rename-term`, `outline`, `refs`.

## Testing

The engine is a pure function (regex over the projection + scope filtering + structural
traversal over a parsed doc dict) → unit-tested with fake docs per repo convention. Docs I/O
stays a thin shell. The rules file is data, exercised via fixtures. Tests use fake IDs/data
only.
