# Claude Code Project Guidelines

This document provides instructions for AI agents (and developers) working on the `python_docx_redline` project.

## Project Overview

A high-level Python API for editing Word documents with tracked changes, eliminating the need to write raw OOXML XML.

**Goal**: Reduce surgical Word document edits from 30+ lines of raw XML manipulation to 3 lines of high-level API calls.

## Development Environment Setup

### Using uv for Dependency Management

This project uses [uv](https://github.com/astral-sh/uv) for fast, reliable Python package management.

#### Initial Setup

```bash
# Install uv if not already installed
curl -LsSf https://astral.sh/uv/install.sh | sh

# Create virtual environment
uv venv

# Activate virtual environment
source .venv/bin/activate  # On macOS/Linux
# .venv\Scripts\activate   # On Windows

# Install package with dev dependencies
uv pip install -e ".[dev]"
```

#### Running Commands in the Virtual Environment

Always run commands from within the activated virtual environment, or use the `.venv/bin/` prefix:

```bash
# Option 1: Activate venv first (recommended for interactive work)
source .venv/bin/activate
pytest tests/ -v
ruff check .
mypy src/

# Option 2: Use direct paths (good for scripts/CI)
.venv/bin/pytest tests/ -v
.venv/bin/ruff check .
.venv/bin/mypy src/
```

#### Managing Dependencies

```bash
# Add a new dependency to pyproject.toml, then:
uv pip install -e ".[dev]"

# Update all dependencies
uv pip install --upgrade -e ".[dev]"

# Sync dependencies exactly as specified
uv pip sync
```

## Code Quality Standards

We maintain high code quality through automated checks. **All code must pass these checks before committing.**

### Required Checks

1. **Ruff** - Linting and code formatting
2. **Pytest** - Unit tests with ≥80% coverage
3. **Mypy** - Type checking (when configured)

### Running Quality Checks

```bash
# Format code with ruff
ruff format .

# Lint code with ruff
ruff check . --fix

# Run tests with coverage
pytest tests/ -v --cov=src/python_docx_redline --cov-report=term-missing

# Ensure minimum 80% coverage
pytest tests/ --cov=src/python_docx_redline --cov-fail-under=80

# Type check (when mypy is configured)
mypy src/
```

### Pre-commit Hooks

We use pre-commit hooks to automatically run quality checks before each commit.

```bash
# Install pre-commit hooks (one-time setup)
pre-commit install

# Run hooks manually on all files
pre-commit run --all-files

# Run hooks on staged files
pre-commit run
```

If hooks fail, fix the issues and re-stage your changes before committing.

## Testing Guidelines

### Writing Tests

- Use **pytest** for all tests
- Place tests in `tests/` directory
- Name test files `test_*.py`
- Name test functions `test_*`
- Use descriptive test names that explain what is being tested

### Test Structure

```python
def test_insert_tracked_basic():
    """Test basic tracked insertion."""
    # Arrange - set up test data
    doc = Document("test.docx")

    # Act - perform the operation
    doc.insert_tracked("new text", after="anchor")

    # Assert - verify the result
    assert doc.has_tracked_changes()
```

### Running Tests

```bash
# Run all tests
pytest tests/ -v

# Run specific test file
pytest tests/test_document.py -v

# Run specific test
pytest tests/test_document.py::test_insert_tracked_basic -v

# Run with coverage
pytest tests/ -v --cov=src/python_docx_redline --cov-report=term-missing

# Run and show coverage in HTML
pytest tests/ --cov=src/python_docx_redline --cov-report=html
open htmlcov/index.html
```

### Coverage Requirements

- **Minimum coverage**: 80%
- Tests must cover core functionality and edge cases
- Aim for 100% coverage on critical paths

## Code Style Guidelines

### Python Style

- **Python version**: 3.10+
- **Line length**: 100 characters (configured in pyproject.toml)
- **Formatting**: Automated by `ruff format`
- **Type hints**: Required for all public functions and methods
- **Docstrings**: Required for all public classes and functions

### Docstring Format

```python
def insert_tracked(
    self, text: str, after: str, author: str | None = None
) -> None:
    """Insert text with tracked changes after a specific location.

    This method searches for the 'after' text in the document and inserts
    the new text immediately after it as a tracked insertion.

    Args:
        text: The text to insert
        after: The text to search for as the insertion point
        author: Optional author override (uses document author if None)

    Raises:
        TextNotFoundError: If the 'after' text is not found
        AmbiguousTextError: If multiple occurrences of 'after' text are found

    Example:
        >>> doc.insert_tracked("new clause", after="Section 2.1")
    """
```

### Import Organization

Organize imports in this order:
1. Standard library imports
2. Third-party imports
3. Local application imports

```python
import tempfile
from pathlib import Path
from typing import Any

from lxml import etree

from .errors import TextNotFoundError
from .text_search import TextSearch
```

## Git Workflow

### Commit Messages

Format:
```
Brief summary in imperative mood (50 chars max)

Longer explanation if needed. Wrap at 72 characters.
Explain what and why, not how.

Key changes:
- Feature 1
- Feature 2

🤖 Generated with [Claude Code](https://claude.com/claude-code)

Co-Authored-By: Claude Sonnet 4.5 <noreply@anthropic.com>
```

### Branch Workflow

- **main**: Stable code, all tests pass
- Feature branches: Created from main for new work
- Always ensure tests pass before merging to main

## Beads Task Management

This project uses [beads](https://github.com/bsilverthorn/beads) for issue tracking.

### Common Commands

```bash
# See all ready work
bd ready

# Claim a task
bd update <issue-id> -s in_progress

# View task details
bd show <issue-id>

# Add a comment
bd comment <issue-id> "Your comment here"

# Close a task
bd close <issue-id> -r "completed"

# See blocked tasks
bd blocked
```

### Workflow

1. **Find work**: `bd ready`
2. **Claim task**: `bd update <id> -s in_progress`
3. **Read details**: `bd show <id>`
4. **Do the work**: Implement according to acceptance criteria
5. **Run quality checks**: `pre-commit run --all-files`
6. **Test**: `pytest tests/ -v --cov-fail-under=80`
7. **Commit**: Git commit with descriptive message
8. **Document**: `bd comment <id> "What was done"`
9. **Close**: `bd close <id> -r "completed"`

## Project Structure

```
python_docx_redline/
├── .venv/                  # Virtual environment (created by uv)
├── src/
│   └── python_docx_redline/      # Main package
│       ├── __init__.py
│       ├── document.py    # Document class
│       ├── text_search.py # Text search algorithm
│       ├── tracked_xml.py # XML generation
│       └── errors.py      # Custom exceptions
├── tests/                 # Test suite
│   ├── __init__.py
│   └── test_document.py
├── docs/                  # Documentation
│   ├── PROPOSED_API.md
│   ├── IMPLEMENTATION_NOTES.md
│   ├── ERIC_WHITE_ALGORITHM.md
│   └── QUICK_REFERENCE.md
├── examples/              # Example YAML files
├── pyproject.toml         # Project configuration
├── CLAUDE.md             # This file
└── README.md             # User-facing documentation
```

## Key Resources

- **API Spec**: `docs/PROPOSED_API.md` - Complete API documentation
- **Implementation Guide**: `docs/IMPLEMENTATION_NOTES.md` - Technical details
- **Eric White's Algorithm**: `docs/ERIC_WHITE_ALGORITHM.md` - Text replacement algorithms
- **Development Plan**: `DEVELOPMENT_PLAN.md` - Project roadmap
- **Examples**: `examples/surgical_edits.yaml` - Real-world use cases

## Common Tasks

### Adding a New Feature

1. Check if there's a beads issue for it: `bd ready`
2. If not, create one: `bd create "Feature description" -t task -p 2`
3. Claim it: `bd update <id> -s in_progress`
4. Create tests first (TDD approach)
5. Implement the feature
6. Ensure all quality checks pass
7. Commit with descriptive message
8. Close the beads issue

### Fixing a Bug

1. Write a failing test that reproduces the bug
2. Fix the bug
3. Ensure the test passes
4. Ensure all other tests still pass
5. Commit the fix

### Updating Dependencies

```bash
# Update a specific package
uv pip install --upgrade package-name

# Update all packages
uv pip install --upgrade -e ".[dev]"

# Verify everything still works
pytest tests/ -v
```

## Troubleshooting

### Tests Failing Due to Coverage Plugin

If you see `error: unrecognized arguments: --cov=src/python_docx_redline`:

```bash
# Ensure pytest-cov is installed
uv pip install pytest-cov

# Or run without coverage temporarily
pytest tests/ -v -p no:cov
```

### Virtual Environment Issues

```bash
# Remove and recreate venv
rm -rf .venv
uv venv
uv pip install -e ".[dev]"
```

### Pre-commit Hooks Not Running

```bash
# Reinstall hooks
pre-commit uninstall
pre-commit install

# Run manually to debug
pre-commit run --all-files
```

## Questions?

- Check `docs/` for technical documentation
- Review `DEVELOPMENT_PLAN.md` for project roadmap
- Search existing beads issues: `bd list`
