# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/). This project is intended to adhere to [Semantic Versioning](https://semver.org/spec/v2.0.0.html) starting with version 1.0.

## [Unreleased]

### Added
### Changed
### Deprecated
### Removed
### Fixed
### Security

## [0.3.0] - 2020-12-03

This release addresses usability issues:

* Previously, `Project`'s associated functions mostly required `&mut self` access. This is a result of exposing implementation details. This requirement has been relaxed, and `Project`'s entire public API requires `&self` access only.
* The whole `ProjectInformation` dance on `Project::information()?` has been flattened into the `Project` struct.

  This makes code like `proj.information()?.information` a thing of the past, at the expense of slightly more parsing during `open_project`. This is entirely acceptable, given that a parse error now is no worse than a parse error later.

### Added

Internal improvements:

* Set up continuous integration workflow.
* Added tests to verify some finicky decompressor internals.

### Changed

* Moved `Project::information()?` to `Project::information`.
* Removed `mut` from `Project`'s public API.

## [0.2.0] - 2020-11-30

### Changed

* Changed `Module::text_offset` from `u32` to `usize`.

### Removed

Breaking changes:

* `Module`:
  * `name_unicode`
  * `stream_name_unicode`
  * `doc_string_unicode`
  * `cookie`

Non-breaking changes:

* `Information`
  * `doc_string_unicode`
  * `help_file_2`
  * `constants_unicode`
* `ReferenceProject`
  * `name.1`
* `ReferenceRegistered`
  * `name.1`
* `ReferenceOriginal`
  * `name.1`
* `ReferenceControl`
  * `name.1`
  * `name_extended.1`

## [0.1.0] - 2020-11-29

### Added

- VBA project parser.
- RLE decompressor for compressed streams.

[Unreleased]: https://github.com/tim-weis/ovba/compare/0.3.0...HEAD
[0.3.0]: https://github.com/tim-weis/ovba/compare/0.2.0...0.3.0
[0.2.0]: https://github.com/tim-weis/ovba/compare/0.1.0...0.2.0
[0.1.0]: https://github.com/tim-weis/ovba/compare/827d416...0.1.0
