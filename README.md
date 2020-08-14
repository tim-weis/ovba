# ovba

Command line utility to inspect and extract Office VBA projects from [Office Open XML](http://www.ecma-international.org/publications/standards/Ecma-376.htm) documents.

## Usage

The *ovba* tool reads from files or standard input (`stdin`), and writes to files or standard output (`stdout`). Errors are reported to standard error (`stderr`) so as to not interfere with output `stdout`.

The following subcommands are available:

* `dump`

  Extract the raw binary VBA project from document. The project file is in the [Compound File Binary File Format](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b).

## License

See the [LICENSE](LICENSE) file for license rights and limitations (MIT).
