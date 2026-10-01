# Repository Guidelines

## Operational constraints

- Verify SCDP pages through the local Playwright connection to the visible Chrome session.
- When login is needed, click **Entrar com gov.br** and load `USERNAME` and `PASSWORD` from `.env`; leave CAPTCHA and other authentication challenges to the operator.
- Keep the project Chrome clone free of extensions; extensions can interfere with gov.br login controls.
- Use `uv run python -m scdp_automation --login` to authenticate without starting extraction.
- Do not commit credentials, browser profiles, generated results, or private travel inputs. Review changes under `input/` carefully.

## Development and validation

Run from the repository root:

- `uv sync` to sync the environment after dependency changes.
- `uv run python -m unittest discover -v` to run tests.
- `uv run ruff check .` and `uv run ruff format --check .` for lint and formatting checks.
- `uv run ty check` for static type checking.

Use Ruff formatting, resolve lint and typing errors directly, and do not add `# noqa` suppressions.
