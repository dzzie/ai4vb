# Answer Key — `testdata/`

Ground truth for the 10 documents, for scoring the `category` column and
sanity-checking `priority` / `sentiment`. The two edge cases (09, 10) are the
interesting ones — that's where models and prompts diverge.

| # | File | category | priority | sentiment | notes |
|---|------|----------|----------|-----------|-------|
| 01 | `01_double_charge.txt` | billing | high | negative | clear billing error + refund + bank-dispute threat |
| 02 | `02_api_outage.txt` | technical | high | negative | production outage, revenue impact, "now not a ticket" |
| 03 | `03_password_reset.txt` | account | low | neutral | login/reset, explicitly "not urgent" — tests low priority |
| 04 | `04_enterprise_demo.txt` | sales | low | positive | inbound prospect, demo + quote, "no rush" |
| 05 | `05_feature_praise.txt` | feedback | low | positive | praise + a feature suggestion |
| 06 | `06_data_deletion.txt` | legal | high | neutral | GDPR Art. 17/15 — formal, no emotion → neutral despite "high" |
| 07 | `07_cosmetic_bug.txt` | technical | low | neutral | cosmetic UI bug, "does not break anything" |
| 08 | `08_invoice_tax.txt` | billing | medium | neutral | invoice reissue / VAT, calm, soft deadline |
| 09 | `09_offsite_memo.txt` | **other** | low | neutral | **off-schema:** an internal offsite memo, fits no support category. Should land `other`. A model that *invents* a fit (e.g. "feedback") is fabricating — the anti-hallucination test. |
| 10 | `10_cancellation.txt` | **account** | high | negative | **ambiguous:** churn/cancellation. Defensible as `account` (cancel request), `billing` (stop billing), or `feedback` (complaint). I scored it `account` because the explicit ask is "cancel my subscription," but disagreement here is *expected and informative* — it's where qwen3 vs Gemini will split. |

## Coverage

- **Categories:** all seven enum values appear — billing (01, 08), technical
  (02, 07), account (03, 10), sales (04), feedback (05), legal (06), other (09).
- **Priorities:** high (01, 02, 06, 10), medium (08), low (03, 04, 05, 07, 09).
- **Sentiments:** negative (01, 02, 10), positive (04, 05), neutral (03, 06, 07, 08, 09).

## How to use it

1. `cat.Scan App.Path & "\testdata", False`
2. `cat.ClassifyExtract ai`
3. Compare the `category` column to this table for accuracy; watch
   `CoercedCount` (should be ~0 — every value here is in-range) and
   `UnparseableCount` (should be 0 if the model returns clean JSON).
4. Run it twice — once with `COllama`/qwen3, once with `CGemini` — and diff the
   two `docs` tables. Rows 09 and 10 are where you'll learn the most.

## Things to watch (the interesting failure modes)

- **09 fabricated into a real category** → the model is forcing a fit rather than
  using `other`. Tighten the prompt's "if nothing fits, use other" clause.
- **06 scored `negative`** → the model is reading formal/legal firmness as anger.
  Sentiment should be `neutral`; a miss here is a sentiment-calibration tell.
- **`valid = 0` on anything** → an enum field came back outside its set (e.g.
  "urgent" instead of "high", or a two-word category). That's a prompt-vs-model
  conformance issue, surfaced by the coercion path — exactly what that column is for.
