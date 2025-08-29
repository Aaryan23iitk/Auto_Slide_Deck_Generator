# Auto Slide Deck Generator

**Author:** Aaryan

A small, practical Python utility that builds a clean, presentation-ready PowerPoint deck on any topic by combining real-time web snippets (DuckDuckGo) with an OpenAI chat LLM. The script produces a strict JSON slide schema from the LLM and converts it into a 7-slide `.pptx` file using `python-pptx`.

---

## Features (each explained)

### 1. Topic-driven slide generation

Provides a simple prompt interface (or `--topic` CLI flag) to tell the script what subject to create slides for. The LLM synthesizes a seven-slide structure: title, overview, four themed content slides (3–6), and conclusion — ready for presentation.

### 2. Real-time web context (DuckDuckGo)

Fetches recent web snippets with `ddgs` (DuckDuckGo search) and builds a compact context block for the LLM. This helps the model include up-to-date facts and conservatively cite trends found in current web content.

### 3. LLM-based JSON output (strict schema)

Asks the OpenAI model to return **STRICT JSON** following a small schema: `slides -> [{title, bullets, notes}]`. Keeping output structured makes validation and generation deterministic and simpler to automate.

### 4. JSON validation + light repair

Parses and validates the returned JSON. If the structure is slightly malformed, a light repair attempt converts obvious variants into the expected schema before creating slides.

### 5. Robust LLM calls with retries

Uses `tenacity` to add exponential-backoff retries around the LLM call, reducing failures from transient network or rate-limit hiccups.

### 6. PPTX generation with styled templates

Uses `python-pptx` to create a simple, clean deck:

* Title slide with large heading, subtitle, and accent line.
* Content slides with navy headers, bullet lists, emoji bullets, and randomized soft backgrounds.
* Presenter notes support.

### 7. CLI-first workflow + dry-run mode

Command-line interface (flags for `--topic`, `--model`, `--max-results`, `--outfile`, `--dry-run`, `--no-web`). Dry-run prints validated JSON instead of writing a file so you can inspect the LLM output first.

### 8. Filename safety & small helpers

Sanitizes output filenames and clamps long web-contexts to control token costs. Small utilities keep produced filenames safe for most filesystems.

---

## What OpenAI key is used

* The script expects an environment variable named `OPENAI_API_KEY` containing your OpenAI API key (the `sk-...` secret). It reads this key at runtime and fails early if not present.
* Optional environment variable `OAI_MODEL` can override the CLI default model (default in code: `gpt-4o-mini`).

---

## Libraries & Dependencies

Minimum Python: **3.10+**

Install the packages with pip:

```bash
pip install -U ddgs python-pptx openai tiktoken tenacity
```

**Primary packages explained (1–2 lines each):**

* `ddgs` — lightweight DuckDuckGo search wrapper used to fetch recent web snippets without requiring API keys.
* `openai` — official OpenAI Python client used to call chat completions and request JSON-formatted content.
* `python-pptx` — used to programmatically create and style PowerPoint `.pptx` files.
* `tiktoken` — (optional) useful if you later want to estimate/trim token usage when building prompts.
* `tenacity` — provides retry decorators (exponential-backoff) for resilient LLM calls.

You can also create a `requirements.txt` with those dependencies for reproducible installs.

---

## Quickstart

1. Install Python 3.10+ and the dependencies above.
2. Set your OpenAI API key in the environment:

   * macOS / Linux: `export OPENAI_API_KEY="sk-..."`
   * Windows CMD: `set OPENAI_API_KEY=sk-...`
   * PowerShell: `$Env:OPENAI_API_KEY = "sk-..."`
3. Run the script:

```bash
python auto_slide_deck.py --topic "Climate Change" --outfile "Climate_Change.pptx"
```

Dry-run example (prints JSON, no PPTX):

```bash
python auto_slide_deck.py --topic "Quantum Computing" --dry-run
```

Skip web scraping (LLM-only):

```bash
python auto_slide_deck.py --topic "AI Ethics" --no-web
```

---

## Internals & Code Structure (short)

* `get_web_results()` — queries DuckDuckGo and builds a compact context trimmed to a max size to control token cost.
* `format_user_prompt()` — builds the user prompt for the LLM, embedding the web context and the required slide schema.
* `call_openai_json()` — talks to the OpenAI client and requests JSON output. Wrapped with retry logic.
* `validate_slide_json()` — enforces a predictable slide list and normalizes bullets and notes.
* `create_pptx()` — uses `python-pptx` to turn validated slide data into a styled PowerPoint file. Adds backgrounds, emoji bullets, and presenter notes.
* `build_deck()` — orchestration function chaining search → prompt → LLM → validation → PPTX creation.

---

## Suggested model & prompt tips

* Production: prefer a low-temperature, instruction-following model (the default `temperature=0.3` works well here).
* If you want longer slide bodies, you can increase the model token limit by selecting a larger model and adjusting the prompt.
* Keep the `LLM_SYSTEM_PROMPT` strict (it instructs the model to return JSON only). The code relies on that structure for automation.

---

## Troubleshooting

* `RuntimeError: duckduckgo-search is not installed.` — install `ddgs` or run with `--no-web`.
* `OPENAI_API_KEY environment variable not set.` — export/set your API key.
* Malformed JSON from LLM — try `--dry-run` to inspect; adjust model (or temperature) and re-run. The script attempts a light repair but complex deviations may require prompt tuning.
* `python-pptx` missing or errors — ensure the package is installed and you are running a supported Python version.

---

## Extensibility ideas

* Add slides with images: fetch royalty-free images and add them to slides.
* Support multiple deck sizes (5/10 slides) through a CLI parameter.
* Add token budgeting with `tiktoken` and truncate web context more intelligently.
* Support exporting speaker notes as a separate text file or PDF.

---

## License & Contribution

Feel free to use and adapt this script for personal or research use. If you want to contribute:

* Open an issue describing your idea or bug.
* Submit a PR with tests and a short description of the change.

---
