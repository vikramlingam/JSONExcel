# Contributing

Thank you for contributing to `jsonexcel`. Keep changes focused, deterministic, and compatible
with the supported Python versions.

## Development setup

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -e '.[dev,docs]'
```

On Windows, activate the environment with `.venv\Scripts\activate`.

## Before opening a pull request

Run the quality checks before opening a pull request:

```bash
ruff check .
mypy src
python -m build
twine check dist/*
mkdocs build --strict
```

Pull requests should:

- explain the problem and user-visible behavior;
- update README or guide documentation when behavior changes;
- avoid generated workbooks, caches, build output, and unrelated formatting changes;
- preserve formula safety, Excel-limit checks, and fail-safe long-text behavior.

Do not add a public feature without a working implementation and usage documentation.
