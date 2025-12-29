# Changelog

All notable changes to python-docx-redline will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2024-12-28

### Added
- **Untracked (silent) editing mode**: New generic editing methods with `track` parameter:
  - `insert()` - Insert text with optional tracking (default: `track=False`)
  - `delete()` - Delete text with optional tracking (default: `track=False`)
  - `replace()` - Replace text with optional tracking (default: `track=False`)
  - `move()` - Move text with optional tracking (default: `track=False`)
- `include_deleted` parameter for `find_all()` to control whether deleted text is searched
  - Default is `False` (excludes deleted text from search results)
  - Use `include_deleted=True` to search within tracked deletions
- Per-edit `track` field support in batch operations (`apply_edits()`)
- `default_track` parameter for batch operations to set default tracking behavior
- Track support for header/footer operations
- Track support for CriticMarkup operations (`apply_criticmarkup()`)

### Changed
- `insert_tracked()`, `delete_tracked()`, `replace_tracked()`, `move_tracked()`
  are now convenience aliases for the generic methods with `track=True`
- `find_all()` now excludes deleted text by default for more intuitive behavior
  - Previous behavior: included deleted text
  - New behavior: excludes deleted text (use `include_deleted=True` for old behavior)
- Updated project description to reflect broader scope beyond just tracked changes

### Deprecated
- None

### Removed
- None

### Fixed
- None

### Security
- None

## [0.1.0] - 2024-12-01

### Added
- Initial release with tracked changes support
- `insert_tracked()`, `delete_tracked()`, `replace_tracked()`, `move_tracked()` methods
- Smart text search that handles run-fragmented text
- Regex support with capture groups
- Fuzzy matching for OCR'd documents
- Quote normalization (curly/straight quote matching)
- Batch operations from YAML/JSON files
- Scope filtering for targeted edits
- Document rendering to PNG
- Image insertion with tracked changes
- Comment support
- CriticMarkup import/export
- python-docx interoperability
