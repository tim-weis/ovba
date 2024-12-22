# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/). This project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html). 0.X.Y releases attempt to communicate breaking changes by a bump in the 0.X version number, while non-breaking changes increment the Y version number only, leaving the X version number unchanged.

## [Unreleased]

### Added
### Changed
### Deprecated
### Removed
### Fixed
### Security

## [0.7.1] - 2024-12-22

Servicing release.

### Changed

* Updated the `cfb` dependency from "0.3" to "0.10".

### Security

* Updated the `nom` dependency from "5.1" to "7.1". This acknowledges [RUSTSEC-2023-0086](https://rustsec.org/advisories/RUSTSEC-2023-0086.html) that describes soundness issues in the `lexical-core` crate.

  `lexical-core` was an optional dependency of the `nom` crate prior to version "7". It was pulled in via default features, but wasn't used by this crate. This update makes sure that clients of this crate won't have to deal with false alarms going forward.

## [0.7.0] - 2024-12-22

### Fixed

* Path separators are now properly handled for non-Windows targets.

## [0.6.0] - 2024-12-19

### Fixed

* The parser acknowledges that the `PROJECTCONSTANTS` record in the `PROJECTINFORMATION` record is optional. It used to assume that it was mandatory.
* The parser acknowledges the presence of an optional `PROJECTCOMPATVERSION` record. This record was first described in version 11 of the \[MS-OVBA\] specification.

## [0.5.0] - 2024-12-02

### Added

* `Information::code_page` is now public.

## [0.4.1] - 2024-01-04

Servicing release.

### Fixed

* Bumps `nom` dependency to at least `5.1.3` to resolve a future incompatibility warning (see PR [Fix for trailing semicolon in macros.](https://github.com/rust-bakery/nom/pull/1657) for details).

## [0.4.0] - 2020-12-07

This release primarily introduces convenience implementations. Additions and removals of `error::Error` variants make this a breaking change.

### Added

* `Project::module_source_raw()`: Convenience implementation to return a module's source code (raw codepage encoding).
* `Project::module_source()`: Convenience implementation to return a module's source code converted to UTF-8 encoding.
* `error::Error::ModuleNotFound`. This is used for public functions that identify modules by name.

### Changed

* Restructured documentation to list module source extraction examples first.

### Removed

* `error::Error::Unknown`: This used to be the catch-all variant. This has been a just-in-case fallback that turned out to never have been used any. It's gone now.

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

[Unreleased]: https://github.com/tim-weis/ovba/compare/0.7.1...HEAD
[0.7.1]: https://github.com/tim-weis/ovba/compare/0.7.0...0.7.1
[0.7.0]: https://github.com/tim-weis/ovba/compare/0.6.0...0.7.0
[0.6.0]: https://github.com/tim-weis/ovba/compare/0.5.0...0.6.0
[0.5.0]: https://github.com/tim-weis/ovba/compare/0.4.1...0.5.0
[0.4.1]: https://github.com/tim-weis/ovba/compare/0.4.0...0.4.1
[0.4.0]: https://github.com/tim-weis/ovba/compare/0.3.0...0.4.0
[0.3.0]: https://github.com/tim-weis/ovba/compare/0.2.0...0.3.0
[0.2.0]: https://github.com/tim-weis/ovba/compare/0.1.0...0.2.0
[0.1.0]: https://github.com/tim-weis/ovba/compare/827d416...0.1.0
