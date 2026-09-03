# MathType sidecar provenance

- Upstream: https://github.com/a917470154/mathtypejx
- Locked commit: `7d90e7274c85cf56ac28d4d15e593044693d7e70`
- Upstream version: `0.1.0`
- License: MIT (see `vendor/MATHTYPEJX-LICENSE.txt` and `vendor/MATHTYPEJX-NOTICE.txt`)
- Runtime: VisualTeX private CPython 3.12.10 x64 archive
- Dependencies: `lxml==6.1.3`, `olefile==0.47`

VisualTeX carries a local validator fix that counts OMML delimiters stored by
`m:d`, including default parentheses and explicit `m:begChr`/`m:endChr`.
Regression tests cover default parentheses, one-sided braces and paired bars.
