# Third-Party Notices

RAS (Better VBA) for Visual Studio Code (`ras-vscode`) is distributed under the MIT License (see `LICENSE`).
It incorporates the third-party material listed below. The original copyright and licence notices are reproduced here as their licences require.

この拡張機能に含まれる第三者の著作物と、そのライセンス・変更点の一覧です。

---

## 1. TypeScript TextMate grammar (microsoft/TypeScript-TmLanguage)

`syntaxes/ras.tmLanguage.json` is a derivative work of the TypeScript TextMate grammar.

| | |
|---|---|
| Upstream project | <https://github.com/microsoft/TypeScript-TmLanguage> |
| Upstream file | `TypeScript.tmLanguage` |
| Upstream version | commit `48f608692aa6d6ad7bd65b478187906c798234a8` (identical to the `master` branch on 2026-11-18; the same commit VS Code's built-in TypeScript grammar is based on) |
| Obtained from | `https://raw.githubusercontent.com/microsoft/TypeScript-TmLanguage/48f608692aa6d6ad7bd65b478187906c798234a8/TypeScript.tmLanguage` |
| SHA-256 of the upstream file | `5112e8f1a32325575e32d371cd4f1b8c4831f4b416147fd80e174b0fa7467566` |
| Unmodified copy in this repository | `third_party/TypeScript-TmLanguage/TypeScript.tmLanguage` (with its `LICENSE.txt`; not shipped in the `.vsix`) |
| Licence | MIT |

### Modifications

The derived grammar is generated from the unmodified upstream file by `scripts/build-grammar.mjs`; that script is the authoritative record of every change. In summary:

1. Converted from property-list XML to JSON. The grammar is named `RAS`, its `scopeName` is `source.ras`, its `fileTypes` are `["ras"]`, the `uuid` is dropped, and a header (`information_for_contributors`, `version`) is added.
2. Every scope name ending in `.ts` is renamed to end in `.ras`.
3. `byref` / `byval` are accepted as parameter modifiers (scope `storage.modifier.ras`) in the rules `parameter-name`, `expression-inside-possibly-arrow-parens` and `type-paren-or-function-parameters`.
4. `inline` / `noinline` are accepted as modifiers in front of `function` in the rule `function-declaration` (scope `storage.modifier.ras`; the capture groups are renumbered).
5. The statement keyword `resume` is added to the rule `control-statement` (scope `keyword.control.ras`).
6. The rule `type-primitive` lists the RAS primitive types (`byte short int long float double currency decimal ptr date string boolean any object void wstr astr`) instead of the TypeScript ones; `number`, `bigint`, `symbol`, `never` and `unknown` are no longer primitives. `object` is removed from the rule `type-builtin-literals`.
7. The rule `numeric-literal` recognises the RAS type suffixes instead of the bigint suffix `n`: `L` / `l` on integer, hexadecimal, binary and octal literals, and `f` / `F`, `d` / `D`, `m` / `M` on decimal literals. The suffix is scoped `storage.type.numeric.ras`.
8. In the rule `decorator`, an unqualified decorator name directly after `@` is scoped `entity.name.function.decorator.ras`.
9. A new rule `struct-declaration` (`declare struct Name { ... }`, scopes `meta.struct.ras`, `storage.type.struct.ras`, `entity.name.type.struct.ras`) is added and included from the rule `declaration`.
10. The rule `support-objects` (names of the JavaScript, DOM and Node.js runtime libraries) and every reference to it are removed.

### Licence

```
Copyright (c) Microsoft Corporation
All rights reserved.

MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
```

---

## 2. Visual Studio Code — TypeScript language configuration (microsoft/vscode)

`language-configuration.json` is adapted from `extensions/typescript-basics/language-configuration.json` of <https://github.com/microsoft/vscode> (`main` branch, 2026-11-18), licensed under the MIT License. The `unbalancedBracketScopes` and `tokenTypes` values of the grammar contribution in `package.json` are taken from `extensions/typescript-basics/package.json` of the same project (with the `.ts` scope suffix renamed to `.ras`).

### Modifications

Comments and trailing commas are removed so that the file is strict JSON, the `["$", ""]` surrounding pair is removed, and regular expressions written as plain strings are written as `{ "pattern": ... }` objects. The patterns themselves are unchanged.

### Licence

```
MIT License

Copyright (c) 2015 - present Microsoft Corporation

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 3. npm packages bundled into `dist/extension.js`

`esbuild` bundles the following packages, unmodified, into `dist/extension.js`. The exact versions are recorded in `package-lock.json`. The build (`esbuild.mjs`) fails if a bundled package has no `### <package>` section below.

Development-only tools (`esbuild`, `typescript`, `@vscode/vsce`, `vscode-textmate`, `vscode-oniguruma`, `plist` and the `@types/*` packages) are not part of the distributed extension.

### vscode-languageclient

<https://github.com/microsoft/vscode-languageserver-node> — MIT License — Copyright (c) Microsoft Corporation. Licence text: see "MIT License (Microsoft Corporation)" below.

### vscode-languageserver-protocol

<https://github.com/microsoft/vscode-languageserver-node> — MIT License — Copyright (c) Microsoft Corporation. Licence text: see "MIT License (Microsoft Corporation)" below.

### vscode-languageserver-types

<https://github.com/microsoft/vscode-languageserver-node> — MIT License — Copyright (c) Microsoft Corporation. Licence text: see "MIT License (Microsoft Corporation)" below.

### vscode-jsonrpc

<https://github.com/microsoft/vscode-languageserver-node> — MIT License — Copyright (c) Microsoft Corporation. Licence text: see "MIT License (Microsoft Corporation)" below.

#### MIT License (Microsoft Corporation)

```
Copyright (c) Microsoft Corporation

All rights reserved.

MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
```

### minimatch

<https://github.com/isaacs/minimatch> — ISC License

```
The ISC License

Copyright (c) 2011-2023 Isaac Z. Schlueter and Contributors

Permission to use, copy, modify, and/or distribute this software for any
purpose with or without fee is hereby granted, provided that the above
copyright notice and this permission notice appear in all copies.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR
IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
```

### semver

<https://github.com/npm/node-semver> — ISC License

```
The ISC License

Copyright (c) Isaac Z. Schlueter and Contributors

Permission to use, copy, modify, and/or distribute this software for any
purpose with or without fee is hereby granted, provided that the above
copyright notice and this permission notice appear in all copies.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR
IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
```

### brace-expansion

<https://github.com/juliangruber/brace-expansion> — MIT License

```
MIT License

Copyright (c) 2013 Julian Gruber <julian@juliangruber.com>

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### balanced-match

<https://github.com/juliangruber/balanced-match> — MIT License

```
(MIT)

Copyright (c) 2013 Julian Gruber <julian@juliangruber.com>

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
of the Software, and to permit persons to whom the Software is furnished to do
so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
