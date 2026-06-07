# ai4vb

**Agentic AI for VB6.** A small, swappable multi-provider AI client suite for VB6, plus a set of examples that put a model to work against live things — an object graph, a database, an image, several models at once, a photo library. The unifying idea: give a model direct access to something live and stay out of its way. Cloud or fully local.


Note: each AI class still has a Form1.list1.additem call for debugging, you can just nuke it, ive kept it so far 

---

## Repository layout

The base classes live at the root — the reusable AI client suite, shared by everything. Each example is a peer folder that builds its own host, prompt, and domain on top of them. No example is privileged; they're five different domains driven by the same backends.

```
/                      base classes: the four AI backend clients + CJSON, CLogger, CFileSystem2
/1_automation_example  agent over a live VB6 object graph via the JScript bridge   (see its README)
/2_db_example          agent given dynamic access to a SQLite database through the app
/3_image_example       upload an image, ask a question about it
/4_multiagent_chat     round-robin debate across multiple backends
/5_img_tagging         resumable image-catalog + vision-tagging app
```

---

## What it does

ai4vb is two things at once:

- **A multi-provider AI client suite** — four backend clients behind one interface, swappable without touching the code that uses them.
- **Five worked examples** — different ways to hand a model live access to data and let it answer questions, from reasoning over VB6 objects to cataloging a folder of photos.

Four backends sit behind one interface. Each example wires in whichever it needs:

- **OpenAI** — `/v1/responses`, server-side context via `previous_response_id`.
- **Anthropic** — `/v1/messages`, client-maintained history (their API is stateless).
- **Google Gemini** — native `generateContent`.
- **Ollama** — local `/api/chat`, no API key, nothing leaves the box.

That last one carries weight: with Ollama the entire loop — model, data, images, reasoning — runs on your own hardware. Sensitive structures never touch a cloud service, there's no key in the product, and no service that can be discontinued.

The deepest pattern here — an iterative loop where the model drives a live object graph through a JScript bridge — is demonstrated by `1_automation_example`. Its mechanics (the run loop, the `host.describe` contract, the earned gotchas) live in that folder's own README.

---

## The base classes

The reusable substrate at the root. Everything else is built per example.

| File | Role |
|---|---|
| `COpenAI.cls` | OpenAI client. Async HTTP with poll-loop and cancel support. Uses `previous_response_id` for context. |
| `CClaudeAI.cls` | Anthropic client. Same surface as COpenAI; maintains conversation history in-memory because Anthropic's API is stateless. |
| `CGemini.cls` | Google Gemini client. Native `generateContent`; auth via `x-goog-api-key`, model in the URL path, vision via `inline_data`. |
| `COllama.cls` | Local Ollama client. Native `/api/chat`, keyless, `stream:false`; vision via base64 image array. Localhost or a remote box via the `RemoteIP` property. |
| `CJSON.cls` | JSON parser/path-accessor over ScriptControl. |
| `CLogger.cls` | Append-only logger with timestamps and section headers. Every run produces a re-readable trail. |
| `CFileSystem2.cls` | File/folder helpers (exists, read, enumerate). |

---

## The two layers

**The framework layer (the base classes):** a swappable multi-provider AI client suite plus the utilities every example leans on — JSON parsing, logging, filesystem access. Backend-agnostic, domain-agnostic.

**The example layer:** each folder adds its own host form, prompt, domain, and — where it uses the JScript bridge — its own bridge module (`modProtoGen` in the automation example, `modDbHost` in the db example). This is where domain expertise plugs in.

The framework doesn't try to make the AI smart. It tries to give the AI everything it needs and stay out of the way.

---

## The examples

Five domains, one set of base classes. Each folder is self-contained with its own host form, prompt, and supporting modules.

**`1_automation_example/`** — the original demo and the reference implementation of the core pattern: an agent reasoning over a live VB6 object graph (`CManager` / `CUser` / `CProject`) through a JScript bridge, backed by any of the four clients. The run loop, the introspection contract, and the prompt gotchas are documented in its own README.

**`2_db_example/`** — gives the AI dynamic access to a SQLite database through the app to answer user questions about the data. `modDbHost` exposes the live DB to the bridge (`cSQLiteTable`, `cSQLiteField`, `CJsonParser`); the agent introspects tables and columns and queries through the host rather than walking a hand-built object graph. Proof the "model" can be a data source, not just VB classes.

**`3_image_example/`** — the simplest case. Upload an image and ask a user question about it: image in, answer out, one shot through the backend clients. No agent loop, no bridge — the minimal multimodal path.

**`4_multiagent_chat/`** — `frmAiChatRoom`: several backends in a round-robin debate over one shared transcript, each assigned a fixed role under an anti-convergence prompt, with mid-run user interjection. Where the single-agent loop reasons, this one argues — and one agent catching another's fabricated statistic turned out to be an emergent feature, not a designed one.

**`5_img_tagging/`** — a complete image-catalog application (a direct VB6 app on the base clients — no bridge). `CImageCatalog` recursively scans a folder, MD5-hashes each file (`modHash`), tags it with a vision model (local qwen2.5-vl via Ollama, or Gemini), and stores everything in SQLite (`images.db`), searchable by tag. Two-phase and resumable: a mechanical scan pass, then a separate AI classify pass you can stop and restart over thousands of images. The prompt is tuned to name the primary subject first and refuse invented scene tags — earned from watching the model miss the car and the house and hallucinate "outdoor" onto studio shots.

A longer writeup of the underlying approach is in `Agentic_Coding_Against_Live_Object_Models.pdf`.

---

## Setup

Each example is self-contained. Open its `Project1.vbp` in VB6, set your API keys in the form (one per cloud backend; keys persist to the registry, Ollama needs ip/model), pick a backend, and run. The bridge examples (automation, db) also need a Regen to generate their protos. See each example's own notes for the specifics.

---

## What this is not

- Not a chatbot. The single-agent examples run to completion or to MAX_STAGES; they're not for free-form conversation. (The multi-agent room is the one deliberate exception.)
- Not a code generator. The agent introspects and reasons about live data, not source files.
- Not async-first. The HTTP layer is async with polling, but the agent loops are otherwise synchronous and run on the UI thread.
- Not a typelib reflector. The bridge examples use source-parsed protos rather than `ITypeInfo`, because Standard EXE VB6 projects don't register typelibs and this works the same on Standard EXE, ActiveX EXE, and ActiveX DLL alike.

---

## Status

The framework is solid and has outgrown its original single example. Tested end-to-end across the cloud and local backends on single-task and multi-task prompts. Self-discovery works (the agent finds new classes added to a project without prompt updates). Cancel works. Hermetic per-stage execution works. Logging captures enough for full post-hoc debugging.

The pattern has now been carried — unchanged at the core — onto a live object graph, a database, a vision model, a resumable cataloging pipeline, and a multi-agent room. Which was the whole bet: deliver a live model to the agent and stay out of the way.
