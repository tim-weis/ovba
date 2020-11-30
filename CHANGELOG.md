# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/). This project is intended to adhere to [Semantic Versioning](https://semver.org/spec/v2.0.0.html) starting with version 1.0.

## [Unreleased]

### Added

Internal:

* Continuous Integration.

### Changed
### Deprecated
### Removed
### Fixed
### Security

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

[Unreleased]: https://github.com/tim-weis/ovba/compare/0.2.0...HEAD
[0.2.0]: https://github.com/tim-weis/ovba/compare/0.1.0...0.2.0
[0.1.0]: https://github.com/tim-weis/ovba/compare/827d416...0.1.0
