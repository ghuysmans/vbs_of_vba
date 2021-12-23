# vbs_of_vba

This is an attempt at converting VBA/VB6 source code to VBScript.

## Limitations
Beware of several limitations:
- `Variant`s are promoted; strict types overflow
- fixed-length `String`s are only initialized with spaces
- `String`s are immutable
- `IIf` doesn't exist in VBScript
