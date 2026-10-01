# Repository Guidelines

## Project Structure & Module Organization

This is a small Python application managed with `uv`. The main executable is
`scdp_automation`; project metadata and development dependencies are defined in
`pyproject.toml`, with resolved versions recorded in `uv.lock`. Tests live
under `tests/`. Verify SCDP pages using a local Playwright script connected to
the project's visible Chrome session; inspect the authenticated interface
through that local session.
Click **Entrar com gov.br** automatically when login is needed. Load `USERNAME`
and `PASSWORD` from the local `.env` file, enter them in the visible browser,
and leave CAPTCHA and other authentication challenges to the operator.

## Build, Test, and Development Commands

Run these commands from the repository root:

- `uv sync` — create or update the virtual environment from `pyproject.toml` and `uv.lock`.
- `uv run python -m scdp_automation` — run the visible SCDP extractor.
- `uv run ruff check .` — run lint checks across the repository.
- `uv run ruff format --check .` — verify formatting without changing files.
- `uv run ruff format .` — apply Ruff formatting.
- `uv run ty check` — run static type checking.

Run the test suite with `uv run python -m unittest discover -v`. New tests go
under `tests/`.

## Coding Style & Naming Conventions

Use Python 3.14-compatible syntax, four spaces for indentation, and descriptive `snake_case` names for functions and variables. Use `PascalCase` for classes and uppercase `CONSTANT_CASE` for module constants. Keep functions focused and add type annotations to public functions and newly introduced interfaces. Ruff is the source of truth for linting and formatting; run it before submitting changes.
Resolve type and lint errors directly; do not add `# noqa` suppressions.

## Testing Guidelines

Add regression tests for new behavior and bug fixes, using filenames such as `tests/test_<area>.py` and test functions named `test_<behavior>`. Until a framework is added, validate changes with the application run, `ruff check`, formatting checks, and `ty check`.

## Commit & Pull Request Guidelines

The repository has no commit history yet, so use concise imperative messages, preferably in lowercase (for example, `add input parser`). Pull requests should explain the motivation, summarize the changes, list validation commands run, and include screenshots when modifying rendered HTML or visual assets. Keep unrelated changes out of the same pull request.

## Configuration and Data Safety

Do not commit secrets, local virtual environments, generated caches, or unrelated downloaded data. Preserve the existing `uv.lock` when changing dependencies, and review changes under `input/` carefully because those files are repository inputs rather than generated build output.
