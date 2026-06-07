# ai4vb — db mode

Note: requires the sqllite4vb library: https://github.com/dzzie/sqllite4vb


Same agent loop architecture as ai4vb (object-model), swapped tool surface:
the AI gets a read-only SQL gate instead of a live COM object graph.

## Files

- `Form1.frm` — patched form. UI unchanged from the object-model version;
  bindings and seed swapped.
- `modDbHost.bas` — new module. Schema dump, safe SQL gate, ASCII grid
  formatter, demo seed.
- `prompt.template.txt` — system prompt template with `{SCHEMA}`
  placeholder. Form1 generates `prompt.txt` from this at startup with
  the live schema spliced in.
- `verify.py` — Python harness mirroring the SeedDb / DumpSchema /
  SafeSql / FormatGrid logic. Not part of the VB6 project; just a way
  to eyeball output before running the real thing.

## Wire-up

1. Add `cSQLite.cls`, `cSQLiteResults.cls`, `cSQLiteField.cls`,
   `cSQLiteTable.cls` to the project (existing, unchanged).
2. Add `modDbHost.bas` to the project.
3. Replace `Form1.frm` with the new one.
4. Drop `prompt.template.txt` next to the .vbp.
5. Run. `Form_Load` opens `:memory:`, seeds the db, writes `prompt.txt`.

## Tool surface visible to the agent

```jscript
host.sql("SELECT ...")    // ASCII grid or "ERROR: ..."
host.answer("...")        // final answer to the UI
```

That's it. No `manager`, no proto files, no describe(). The schema is
already in the system prompt.

## Demo questions worth running

Each of these requires the agent to actually *read* the data, not guess
from the schema:

- *"Who logs the most hours on the highest-budget project?"*
  Trap: Carol leads Sentinel ($420k) but does zero hours on it.
  Correct answer: Alice (86 hours).

- *"Are there any tasks without an assignee?"*
  Trap: NULL in a join column — needs LEFT JOIN or IS NULL.

- *"Which project leads aren't doing any tasks on their own project?"*
  Multi-hop self-comparison.

- *"List active projects, their leads, and total hours logged."*
  status='active' filter + 2-way join + aggregate.

## ⚠️ One thing to confirm

I used `db.Execute "..."` in `modDbHost.SeedDb` for DDL and INSERTs.
The cSQLiteTable code only shows `db.Query(...)` being called, so I
inferred a sibling `Execute` method exists on cSQLite for non-query
SQL. If your cSQLite class names this differently (e.g. `Exec`,
`ExecuteNonQuery`, `RunSql`), search/replace `db.Execute` in
`modDbHost.bas` accordingly. Same for `db.Open ":memory:"` in Form1 —
if your open is `.OpenDatabase` or `.Connect`, adjust there.

If cSQLite has no non-query method at all, the seed can be done by
`db.Query`-and-discard, since SQLite happily accepts DDL through the
prepare/step path; just call `Set rs = db.Query "...": rs.Close` (or
similar) for each statement.
