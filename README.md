# xlextract

`xlextract` is an offline Windows-first CLI for extracting trustworthy structured
data from Excel workbooks and publishing the evidence needed to prove the result.

The supported command name is `xlx`. V1 targets Python 3.12 and a portable
Windows 10/11 one-directory distribution. Implementation starts with the CLI and
core; the TUI is optional and late.

## Development checks

```powershell
powershell -File tools/ci.ps1
```

The project does not make network calls at runtime or during its test suite.
