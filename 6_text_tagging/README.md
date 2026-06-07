# 6 — Text → Structured Output

The most common production LLM task: take free-form text in, get structured
fields out. Support tickets → `{category, priority, sentiment}`. Resumes →
`{name, years_exp, skills}`. Log lines → `{severity, component, error_code}`.
This example is the text sibling of `5_img_tagging` — identical two-phase,
resumable skeleton, but the model returns a **JSON object** instead of a tag
list, and the harness **validates** it before trusting it.

## The pattern (how it's commonly done)

1. **Constrain the output in the prompt.** Hand the model an exact schema and
   demand *"ONLY a JSON object, no markdown."* Enum fields list their allowed
   values inline.
2. **Sanitize before parsing.** Models wrap JSON in ```` ```json ```` fences or
   add *"Here's the object:"*. `ExtractJsonObject` grabs the outer `{ … }` and
   throws the rest away — the structured-output cousin of the chatroom's
   `StripSpeakerPrefix`. You never trust the model's framing.
3. **Parse, then validate the constrained fields.** `CJSON` parses; `Coerce`
   checks each enum field against its allowed set and falls back if it's out of
   range. Free-text fields (`summary`, `keywords`) pass through; the enums are
   where models drift, so that's all you police.
4. **Materialize the intermediate.** The raw JSON is stored in `raw_json`
   alongside the parsed columns — so you can re-validate or re-parse later
   *without re-calling the model*, and any bad field is auditable instead of a
   black box. Same principle as keeping the md5, the tags, the transcript.
5. **Resumable two-phase.** `Scan()` (mechanical: stat + MD5, no AI) then
   `ClassifyExtract()` (the AI step, only touches `classified = 0`). Cancel a
   10k-file run, resume tomorrow; a changed file re-queues just that row.

## Files

- `CTextCatalog.cls` — the engine. Drop in next to your `cSQLite` wrapper,
  `CFileSystem2`, `CJSON`, and `modHash` (the `FileMD5` from example 5).

## Flag semantics

| `classified` | meaning |
|---|---|
| `0`  | pending (also where a transient non-200 leaves it → auto-retry) |
| `1`  | extracted OK |
| `-1` | got a 200 but the JSON was unparseable → **needs your eyes**; `raw_json` holds what it produced |

`valid = 0` marks a row where an enum field was coerced to a fallback.

## Eval metrics — built into the schema

Because validation is recorded, the catalog *is* the eval harness:

- `CoercedCount` → schema-conformance miss rate (enum fields out of range).
- `UnparseableCount` → JSON parse-failure rate.
- Category accuracy → label a sample set, compare to the `category` column.

Run the same folder through `COllama` (qwen3, local) and `CGemini`
(`gemini-3.5-flash`) and the numbers tell you whether the gap is the **model**
or the **prompt** — the same A/B/C move from the image example, where the gap
turned out to be mostly prompt.

## Local-first, by design

Defaults to `COllama` + a local text model because this is the workload that
justified the box: **smart categorization over sensitive data that should
never leave it.** Set `ai.Think = False` for batch runs (you want the answer,
not the reasoning trace) and reserve a cloud model for non-sensitive benchmarks.
Swapping providers is one line — every client exposes the same `CreateResponse`.

## Adapting the schema

Everything domain-specific lives at the top of the class:

- `ALLOWED_CATEGORY` / `ALLOWED_PRIORITY` / `ALLOWED_SENTIMENT` — pipe-delimited
  allowed sets. Edit these and the matching `FALLBACK_*` constants.
- `EXTRACT_PROMPT` — restate the same fields and allowed values here so the
  prompt and the validator agree.

Add a field by: extending the prompt, adding a column in `OpenCatalog`, pulling
it with `j.getVal("yourField")`, and (if it's constrained) routing it through
`Coerce`.
