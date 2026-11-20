# RAS 言語仕様書（Draft v0.9）

> v0.8 からの変更
> * Excel / VBA / Office のメンバーとグローバルは、JavaScript の命名規則（camelCase）で参照する（`ws.cells(r, 1).value`、`msgBox(...)`）。生成する VBA にはタイプ ライブラリの綴りを出力する（§15.3.1）
> * `worksheets` などのホストのコレクションで、配列と同じ `filter` / `map` / `forEach` 等を使える（§4.4）
> * `ras check --fix` とエディタのクイックフィックスで、VBA の綴りで書かれた名前を書き換える（§7.3、§16.2）
> * 名前付き引数 `f(a: 123, c: "あいう")` を追加（§3.3）。エディタで引数名を補完する（§16.2.1）
> * ユーザーフォームのデバッグ実行（§14.10.7）と Excel での実行（§14.11.8）を追加
> * 参照の検索・名前の変更の対象と拒否する名前、セマンティックトークンの種類、未インポートの候補の挿入位置を規定（§16.2.5〜§16.2.7）
> * ユニオン型 `A | B` と、`instanceof` / `typeof` による型の絞り込みを追加（§2.6）。同梱の型定義で Object / Variant を返していたメンバーのうち、値の型が決まっているもの（`activeSheet`・`sheets(1)` は `Excel.Worksheet | Excel.Chart`、`getOpenFilename()` は `string | boolean | any[]` など）をユニオン型にした（§2.6.5）
> * クラス・インターフェイスの型パラメータに既定の型引数 `T = 型` を書けるようにした（§2.4.1）

> v0.7 からの変更
> * 構造体（`struct`）をデータの入れ物の基本とし、オブジェクトリテラル・無名のオブジェクト型・タプル・分割代入を追加（§2.5）
> * 値が構造体の `Map<K, V>` は、構造体の配列とキー → 添字の対応を持つクラスに自動変換する（§2.5.6）

> v0.6 からの変更
> * エディタ支援（IntelliSense: メンバー一覧・クイック情報・パラメーター情報等）とドキュメントコメント（JSDoc 形式）を追加（§16）
> * 参考実装と採用ライブラリを選定（§17）。§7.2 の「主要クレート（想定）」を §17 に置き換え
> * 未決事項を §18 に移動

> v0.5 からの変更
> * ジェネリクスを追加。プロジェクト内で実際に使われる型引数ごとに、別の関数・クラスとして出力する（§2.4）

> v0.4 からの変更
> * 長時間ループへの `DoEvents` 自動挿入（間引き付き）を追加（§10.4）。それに伴い、マクロの再入防止と Esc キーでの中断を追加（§12.4）
> * 対象ホストを Excel に限定（Access / Word は当面サポートしない）（§1）

> v0.3 からの変更
> * `number` 型を廃止し、数値型を C# 風の `short` / `int` / `long` / `float` / `double` に変更（§2.1〜2.3）
> * C# 風の拡張メソッドを追加（§3.2）

> v0.2 からの変更
> * 実装言語を Rust に確定（§7.2）
> * Mac 版 Excel は対象外と確定（§1）
> * Excel を起動せずに `.xlsm` 等へ直接書き込むオフライン書き込みを初版から対応（§14.9）
> * Range ⇔ 2次元配列の添字を 0 始まりに統一（§9）

> v0.1 からの主な追加・変更
> * §7 ツールチェインを「シングルバイナリ `ras` ＋ 薄い VS Code 拡張」構成に変更
> * §2 型マッピングを性能重視に変更（型が確定する変数は `Variant` ではなく具体型で出力）
> * §4 ラムダ式を「インライン展開を第一とし、クラス生成はエスケープ時のみ」に変更
> * §5 `finally` / `using` / `throw` と、1プロシージャ内での複数 `try` の変換規則を追加
> * 新設: §8 最適化（`inline`・自動展開）、§9 Range ⇔ 2次元配列の自動一括変換、§10 アプリケーション状態制御、§11 `declare`（32/64bit 両対応）、§12 ボイラープレート自動生成、§13 VBA 罠の自動回避、§14 Excel 同期（ライブ／オフライン）、§15 プロジェクト構成、§16 未決事項

---

## 1. 設計思想・基本方針

* **ターゲット:** 標準VBA（`.bas`, `.cls`、および Excel のドキュメントモジュール `ThisWorkbook` / シートモジュール）へ変換するトランスパイル言語。
* **シンタックス:** TypeScriptのサブセット＋VBA固有仕様を吸収する最小限の拡張。既存のTS系エディタ資産（TextMate Grammar等）を流用しやすい構文を維持する。追加キーワードは `byref` / `byval` / `inline` / `noinline` / `resume` のみとし、それ以外の拡張はデコレータ（`@xxx`）と TS 既存構文（`declare module`, `using` 等）の転用で表現する。
* **静的検証の委譲:** 厳密な型チェックや制約の検証はRAS側（コンパイラ兼LSP）が担う。
* **ボイラープレートゼロ:** VBA で「毎回書くべき定型コード」「知らないと踏む罠の回避コード」は、ユーザーが書かずともコンパイラが自動で挿入する（§12, §13）。
* **速度はコンパイラの責務:** VBA は関数呼び出し・COM 呼び出し（セルアクセス）・`Variant` 演算・遅延バインディングが遅い。RAS ではユーザーが素直に書いたコードを、コンパイラがインライン展開・一括転送・具体型化によって高速な VBA に変換する（§8, §9）。
* **対象環境:** Windows 版 Excel（Office 2007 以降、32bit / 64bit）。**Access / Word 等の他の Office ホストは当面サポートしない。Mac 版 Excel は対象外**とする（Mac 固有の `#If Mac` 分岐・`Lib` パス・AppleScript 連携等は生成しない）。
* **出力 VBA は読める状態を保つ:** 生成コードは人間がデバッグ可能であること。各ステートメントには元の RAS ソース位置を辿れるソースマップを持つ（§14.8）。

---

## 2. 型システムと変数宣言

* **変数・定数宣言:** `let`（再代入可能）および `const`（再代入不可）を使用する。

### 2.1 基本データ型とVBA型の対応（v0.4で変更）

TypeScript の `number` 型は存在しない（`number` と書いた場合は「`int` または `double` を使う」旨のエラー）。数値型は C# の名前に揃える。

| RAS 型 | VBA 型 | サイズ | 備考 |
|---|---|---|---|
| `byte` | `Byte` | 8bit 符号なし | |
| `short` | `Integer` | 16bit | |
| `int` | `Long` | 32bit | 整数リテラルの既定型 |
| `long` | `LongLong` | 64bit | 64bit Office 専用。32bit を含むターゲットでは `#If Win64` ガード外での使用をエラー（§11.3） |
| `float` | `Single` | 32bit 浮動小数点 | |
| `double` | `Double` | 64bit 浮動小数点 | 小数リテラルの既定型 |
| `currency` | `Currency` | 64bit 固定小数点 | |
| `decimal` | `Variant`（`CDec`） | 96bit | VBA に `Decimal` 型の変数宣言は無いため `Variant` に格納 |

| RAS 型 | VBA 型 | 備考 |
|---|---|---|
| `ptr` | `LongPtr`（VBA7）/ `Long`（VBA6） | ポインタ・ハンドル。§11 参照 |
| `currency` | `Currency` | |
| `date` | `Date` | |
| `byte` | `Byte` | |
| `string` | `String` | |
| `boolean` | `Boolean` | |
| `any` | `Variant` | 型不明・異種混在 |
| `object` | `Object` | 遅延バインディング |
| `A \| B`（ユニオン型） | `Object` / `Variant` | メンバーがすべてオブジェクト型なら `Object`、それ以外は `Variant`。§2.6 |
| `Excel.Range` 等 | `Range` 等（事前バインディング） | ホスト型ライブラリ由来。§15.3 |
| RAS の `class` | 生成クラス名 | |
| `struct` / `{ a: A; b: B }` / `[A, B]` | `Type`（ユーザー定義型） | 値型。§2.5 |
| `T[]` | `T()` の動的配列 | §13 の配列規則参照 |

### 2.2 数値リテラル

| 書き方 | 型 | VBA 出力 |
|---|---|---|
| `123` | `int`（`int` に収まらない場合は `long`） | `123&` |
| `123L` | `long` | `123^` |
| `1.5`, `1e3`, `2d` | `double` | `1.5#` |
| `1.5f` | `float` | `1.5!` |
| `1.5m` | `decimal` | `CDec("1.5")` |
| `0xFF` | `int` | `&HFF&` |

* VBA の整数リテラルは値が小さいと `Integer` 型になり、`300 * 200` が `Integer` 演算で溢れる。RAS は**数値リテラルに常に型サフィックスを付けて出力**し、この罠を回避する。
* `short` / `byte` 型の変数への代入は、定数がその型の範囲内であれば暗黙に許可する（C# と同じ）。

### 2.3 数値演算と型変換

C# の規則に従う。

* **二項演算の型:** 両辺の型を次の順で「大きい方」に揃えて演算する: `int` < `long` < `float` < `double`（`currency`, `decimal` は混在時に明示キャストを要求）。`byte` / `short` 同士の演算は `int` に昇格する（VBA の `Integer * Integer` が `Integer` のまま溢れる罠を、`CLng` を挿入して回避する）。
* **整数同士の `/`:** 整数除算（0 方向への切り捨て）。VBA の `\` に変換する（オペランドが整数型なので VBA の丸めは発生しない）。実数除算にしたい場合は片方を `double` にする。
* **`%`:** 整数同士なら VBA の `Mod`（符号は被除数に従い、C# と一致）、実数を含む場合はランタイム `RAS_FMod`。
* **オーバーフロー:** 整数演算の溢れは実行時エラー（VBA のエラー 6）。定数式の溢れはコンパイルエラー（C# の `checked` 相当）。
* **暗黙の型変換:** 拡大変換（`byte`→`short`→`int`→`long`、`int`→`double`、`float`→`double` 等）のみ暗黙に許可する。縮小変換は `x as int` で明示する（TS の `as` 構文を流用）。
* **実数→整数の変換は切り捨て:** VBA の `CLng` は銀行丸めのため、`x as int` は `CLng(Fix(x))` に変換して C# と同じ 0 方向切り捨てにする。丸めたい場合は `Math.round(x)`。

* **具体型出力の原則:** 型が静的に確定する変数・引数・戻り値は具体型で宣言する（`Variant` は約2倍以上遅いケースがあるため）。型が `any` の箇所のみ `Variant` を使う。
* **`Set` の自動解決:** オブジェクト型の代入には型情報から `Set` を自動付与する。`any` 型への代入で右辺の型が静的に決まらない場合は、ランタイムヘルパ `RAS_Assign`（`IsObject` で分岐）を使用する。
* **型推論:** 初期化式の型をそのまま使う（`let x = 1` は `int`、`let y = 1.5` は `double`）。

### 2.4 ジェネリクス

関数・拡張メソッド・クラスに型パラメータを使える。VBA にジェネリクスは無いため、**プロジェクト内で実際に使われている型引数の組み合わせごとに、別の関数・クラスとして出力する**（単相化）。`Variant` 版の共通実装は出力しない。

```typescript
export function maxOf<T extends Comparable>(this: T[]): T {
    let m = this[0];
    for (const v of this) if (v > m) m = v;
    return m;
}

class Stack<T> {
    private items: T[] = [];
    push(x: T): void { this.items.push(x); }
    pop(): T { return this.items.pop(); }
}

const a = ints.maxOf();          // T = int
const b = names.maxOf();         // T = string
const s = new Stack<Excel.Range>();
```

**出力VBAコード:**
```vba
Public Function maxOf__Long(ByRef this_() As Long) As Long
    ' ...
End Function

Public Function maxOf__String(ByRef this_() As String) As String
    ' ...
End Function
```
クラスは `Stack__Range.cls` のように型引数ごとに1つのクラスモジュールを出力する。

#### 2.4.1 規則

* **型引数の決定:** 呼び出し時の型引数は、引数の型から推論する。推論できない場合は `f<int>(...)` のように明示する。決まらなければエラーにする。
* **出力する組み合わせ:** プロジェクト内の全呼び出し（他のジェネリック関数の中からの呼び出しを含む）を推移的にたどり、到達した組み合わせだけを出力する。一度も呼ばれないジェネリック関数は何も出力しない。それが `export` されている場合は情報を表示する。
* **名前:** `元の名前__VBA型名[_VBA型名...]` とする（例: `maxOf__Long`、`Map__String_Long`）。配列は `Arr` を付ける（`T = int[]` なら `__LongArr`）。名前が VBA の識別子長の上限（255 文字）を超える場合は、短縮名とハッシュを使う。
* **同じ VBA 型の共有:** 型引数が違っても、出力される VBA の型表記が同じになる組み合わせ（例: `any` と `decimal`、遅延バインディングの `object` 同士）は1つの実装を共有する。
* **マクロ入口・外部からの呼び出し:** ジェネリック関数は、マクロ入口・`@on`・`Application.Run` の対象にはできない（エラー）。外から呼ばせたい場合は、具体型の非ジェネリック関数で包む。
* **`Set` と比較演算:** `Set` の要否（§2）、`==` を `Is` と `=` のどちらにするか（§13.2）、`Variant` 経由の変換が必要かは、展開後の具体型ごとに決める。各実装は具体型のコードとして最適化され、インライン展開（§8）の対象にもなる。
* **型パラメータの制約:** `T extends 型` で制約を書ける。標準で次の制約を用意する。
  * `Comparable`: `<`, `>` 等が使える型（数値型・`string`・`date`・`currency`）
  * `Numeric`: 算術演算が使える型
  * `ObjectType`: オブジェクト型（`Set` で代入する型）
  
  制約を満たさない型引数はコンパイル時にエラーにする。
* **無限展開の禁止:** 呼び出しのたびに型引数が大きくなる再帰（`f<T>` の中で `f<T[]>` を呼ぶ等）はエラーにする。
* **既定の型引数（v0.9で追加）:** クラス・インターフェイスの型パラメータには `T = 型` で既定の型引数を書ける。型引数を省略して書いた型はその型引数になる（`interface Sheets<T = Worksheet | Chart>` の `Sheets` は `Sheets<Worksheet | Chart>`）。既定の型引数を持つ型パラメータより後ろの型パラメータにも既定の型引数が必要（`E0171`）。関数の型パラメータには書けない（型引数は引数から推論する）。
* **コード量の警告:** 1つのジェネリック定義から出力される実装数が閾値（既定 16）を超えたら警告する。展開後のプロシージャ・モジュールのサイズ制限は §13.8 の規則で検査する。
* **ラムダ:** ジェネリック関数内でクラス化されるラムダ（§4.1）も、型引数の組み合わせごとに別のクラスとして出力する。

### 2.5 構造体・オブジェクトリテラル・タプル・分割代入（v0.8で追加）

複数の値をまとめる入れ物は、VBA の `Type`（ユーザー定義型）に変換する。`Type` はフィールドの位置がコンパイル時に決まり、生成時にメモリ確保も参照カウントも起きないため、クラス（`New` ごとに COM オブジェクトを生成）や `Scripting.Dictionary`（遅延バインディングとハッシュ検索）より軽い。配列にすると要素が連続した領域に並ぶ。

#### 2.5.1 構造体とオブジェクトリテラル

```typescript
struct Point { x: int; y: int; name?: string; }

function make(x: int, y: int): Point {
    return { x, y };              // name は省略（`?` 付きのフィールドだけ省略できる）
}
```

```vba
Public Type Point
    x As Long
    y As Long
    name As String
End Type

Private Function make(ByVal x As Long, ByVal y As Long) As Point
    make.x = x
    make.y = y
End Function
```

* オブジェクトリテラル `{ a: 1, b: x }`（省略記法 `{ x, y }` を含む）は、期待される型が構造体のとき、その構造体の値になる。フィールドの過不足・型の不一致はコンパイル時にエラーにする。`name?: T` と宣言したフィールドだけ省略でき、省略したフィールドは既定値（`0` / `""` / `false` / 空配列）になる。
* リテラルは代入先のフィールドへ直接書き込む（一時変数とコピーを作らない）。`return { ... }` は戻り値へ、`arr.push({ ... })` は配列の要素へ直接書き込む。右辺が代入先を参照している場合（`p = { x: p.y, y: p.x }`）だけ、一時変数に組み立ててから代入する。
* 構造体は**値型**である。代入はコピーで、`const q = p; q.x = 1;` は `p` を変更しない（TypeScript のオブジェクトとは異なる）。関数には参照で渡す（`ByRef`）。

#### 2.5.2 無名のオブジェクト型

期待される型が無いオブジェクトリテラルは、フィールド名と値の型から決まる無名の構造体になる。型注釈にも `{ a: A; b: B }` の形で書ける。フィールドの名前と型が同じなら、書いた順序によらず同じ型である（1 つの `Type` にまとめて出力する）。

```typescript
const o = { count: 1, label: "two" };          // 型は { count: int; label: string }
function stats(xs: double[]): { min: double; max: double } { ... }
```

名前の付いた `struct` とは、フィールドが同じでも別の型である（代入できない）。

#### 2.5.3 タプル

`[A, B]` は位置でフィールドを持つ無名の構造体である。期待される型がタプルのとき、配列リテラルがその値になる。要素は整数リテラルの添字（`t[0]`）で参照する。

```typescript
function divmod(a: int, b: int): [int, int] {
    return [a / b, a % b];
}
```

```vba
Public Type RAS_Tup1
    item0 As Long
    item1 As Long
End Type

Private Function divmod(ByVal a As Long, ByVal b As Long) As RAS_Tup1
    divmod.item0 = a \ b
    divmod.item1 = a Mod b
End Function
```

無名のオブジェクト型とタプルのフィールドに、型パラメータを含む型は使えない（`struct` を宣言する）。

#### 2.5.4 分割代入

`const` / `let` の宣言と `for...of` の変数に、オブジェクトパターン `{ a, b: 別名 }` と配列パターン `[a, , c]` を書ける。構造体・タプルのフィールド、配列の要素を取り出す宣言に展開する。入れ子のパターン・既定値・`...` には対応しない。

```typescript
const { x, y: py } = make(1, 2);
const [q, r] = divmod(17, 5);
for (const { x, y } of points) total += x * y;
```

```vba
ras_d1 = make(1&, 2&)
x = ras_d1.x
py = ras_d1.y
```

#### 2.5.5 構造体の配列

`T[]` の `T` が構造体のとき、`Dim a() As T` の配列になる。ループ内の `push` は容量倍増方式（§13.3）で、リテラルは要素へ直接書き込む。

`for (const p of ps)` は VBA では要素を `p` にコピーするが、ループ本体が `p` のフィールドを読むだけで、配列がループ中に変更され得ない場合は、コピーせず要素を直接参照する。

```typescript
for (const p of ps) total += p.x + p.y;
```

```vba
For ras_i = 0 To RAS_Len_Point(ps) - 1&
    total = total + (ps(ras_i).x + ps(ras_i).y)
Next ras_i
```

直接参照にしない（コピーする）のは、`p` 全体を使う（代入・関数に渡す）・`p` のフィールドを書き換える・ループ本体が配列自体に触れる、のいずれかの場合。配列が引数のときは、さらにループ本体が関数を呼ばず、配列の要素へ代入しないことを条件にする。

構造体の配列は `Variant` に入れられないため、長さはランタイムの `RAS_Len` ではなく、構造体ごとに生成する `RAS_Len_<型名>` で求める。同じ理由で `indexOf` / `includes` / `join` / `sort` / `reverse` は使えない（`findIndex` などコールバックを取るメソッドは使える。コールバックはアロー関数を直接書く）。

#### 2.5.6 値が構造体の `Map`

構造体は `Scripting.Dictionary` に入れられない。`Map<K, V>` の `V` が構造体のとき、コンパイラはこの型を、値を構造体の配列に並べ、Dictionary にはキー → 添字の対応だけを持たせるクラス（`RAS_StructMap__<K>_<V>.cls`、`K` と `V` の組み合わせごとに 1 つ）に置き換える。使い方は通常の `Map` と同じである（`get` / `set` / `has` / `delete` / `size` / `keys()` / `values()` / `clear()`）。

```typescript
const m = new Map<string, Point>();
m.set("a", { x: 1, y: 2 });
total += m.get("a").x;
```

* `get` は値のコピーを返す。登録されていないキーには、全フィールドが既定値の構造体を返す（`has` で確認する）。値を変更するときは、変数に取り出して書き換え、`set` し直す（`m.get(k).x = 1` はコピーへの代入になるためエラーにする）。
* 添字での参照（`m[k]`）には対応しない。
* マップ 1 つにつきオブジェクトは 1 つで、要素ごとのオブジェクト生成は起きない。

#### 2.5.7 構造体を使えない場所

VBA の `Type` は `Variant` に変換できず、クラスモジュールの公開メンバーにも出せない。次はコンパイル時にエラーにする。

* `any` / `object` への代入、`any` を受け取る関数（`console.log` を含む）へ渡すこと。フィールドを個別に渡す。
* クラスの `public` フィールド（`private` にする）、インターフェイスのメンバー。クラスのメソッドの引数・戻り値には使える（`Friend` で出力する）。
* 変数に入れる・引数として渡すアロー関数（クラスに変換されるもの。§4.1）の引数・戻り値・キャプチャ。
* 構造体の 2 次元配列、`null`。

参照を共有したい・`null` を持たせたい・`any` に入れたい場合は `class` を使う。

### 2.6 ユニオン型（v0.9で追加）

いくつかの型のどれかの値を持つ型を `A | B` と書く。`any` と違い、使えるのはすべてのメンバーに共通する操作だけで、それ以外は型ガード（§2.6.4）で型を絞り込んでから使う。

```typescript
function describe(s: Excel.Worksheet | Excel.Chart): string {
    s.activate();                                  // 共通のメンバーは絞り込まずに使える
    if (s instanceof Excel.Worksheet) {
        return `${s.name}: ${s.usedRange.address()}`;  // ここでは s は Excel.Worksheet
    }
    return s.chartTitle.text;                      // ここでは s は Excel.Chart
}

function cellText(v: string | double): string {
    if (typeof v == "String") return v;            // v: string
    return v.toFixed(2);                           // v: double
}
```

**出力VBAコード:**
```vba
Private Function describe(ByVal s As Object) As String
    s.Activate
    If (TypeOf s Is Worksheet) Then
        describe = s.Name & ": " & s.UsedRange.Address()
        Exit Function
    End If
    describe = s.ChartTitle.Text
End Function

Private Function cellText(ByVal v As Variant) As String
    If TypeName(v) = "String" Then
        cellText = v
        Exit Function
    End If
    cellText = RAS_ToFixed(v, 2&)
End Function
```

#### 2.6.1 型と格納先

* **VBA の型:** メンバーがすべてオブジェクト型（クラス・インターフェイス・ホストの型・`object`・`Map`・関数型）なら `Object`、それ以外は `Variant` に格納する。メンバーのアクセスは遅延バインディングになる。
* **正規化:** 入れ子のユニオン型は平らにし、重複したメンバーは除く。メンバーは値型・オブジェクト型の順に並べて表示する（`int | string` は `string | int`）。`any` を含むと `any`、`object` を含むと他のオブジェクト型は `object` にまとまる。
* **`null`:** `| null` はオブジェクト型を含むユニオン型にだけ書ける（`Excel.Range | string | null`）。`string | int | null` はエラー。
* **含められない型:** `void`、構造体（`Variant` に入らないため。§2.5.7）、型パラメータ、`wstr` / `astr`。
* 配列の要素・`Map` の値・型引数・型エイリアスにも使える（`type Shape = Circle | Square;`、`(Circle | Square)[]`、`Box<string | int>`）。

#### 2.6.2 代入と演算

* **代入:** メンバーのどれかに代入できる値は、ユニオン型に代入できる。ユニオン型の値は、すべてのメンバーを代入できる型に代入できる（`Excel.Worksheet | Excel.Chart` は `object` に代入できる）。
* **一部のメンバーだけが代入できる型への代入:** 警告（`W0306`）を出して代入する（`const ws: Excel.Worksheet = activeSheet;`）。VBA の `Set ws = ActiveSheet` と同じく、値が代入できないメンバー（グラフ シート）だったときは実行時エラーになる（値型なら VBA の暗黙の型変換）。`instanceof` / `typeof` で絞り込むか、`as` で変換すると警告は出ない。どのメンバーも代入できない型への代入はエラー。
* **数値の拡大変換:** メンバーの型そのものではなく拡大変換で入る値は、そのメンバーの型に変換して格納する（`string | double` に `1` を入れると `CDbl(1&)`）。`typeof` が格納した値の型名を返すため。
* **`==` / `!=`:** どれかのメンバーと比較できれば書ける。`Variant` に格納するユニオン型の比較では、値の種類（数値・文字列・真偽値・日付・オブジェクト）が違えば等しくない（VBA の `=` のように変換したりエラーにしたりしない。ランタイム `RAS_UnionEq`）。
* **大小比較・文字列の連結:** すべてのメンバーで使える場合だけ書ける（`int | double` どうしの `<`、`"a" + v`）。算術演算は絞り込んでから行う。
* **`as`:** どれかのメンバーとの間で変換できれば書ける（`v as double`）。
* **条件演算子・配列リテラル:** 型の異なる値からユニオン型は推論しない（`cond ? 1 : "a"` はエラー `E0302`）。`const x: string | int = ...` のように型注釈を付ける。

#### 2.6.3 メンバー

* **共通のメンバー:** プロパティ・メソッドは、すべてのメンバーの型に同じ名前のメンバーがあるときに使える。補完にも共通のメンバーだけを出す。値型を含むユニオン型のメンバーは参照できない（レシーバとしてユニオン型を受け取れる拡張メソッドは呼べる）。
* **プロパティ:** 型はメンバーごとの型のユニオン型になる。型がメンバーごとに異なるプロパティ、どれかで読み取り専用のプロパティには代入できない。
* **メソッド:** 引数は、すべてのメンバーのメソッドで受け取れなければならない。戻り値はメンバーごとの戻り値のユニオン型（どれかが `void` なら `void`）。名前付き引数は、どのメソッドでも引数の名前と順序が同じときだけ使える。ジェネリックなメソッドは呼べない。
* **一部の型にだけあるメンバー:** ホストの型（Excel など）のメンバーが一部のメンバー型にしかないときは、警告（`W0305`）を出し、それを持つ型のメンバーとして型を付けて遅延バインディングで呼ぶ（`activeSheet.cells(1, 1)` は `Excel.Range`。グラフ シートがアクティブなら実行時エラー）。RAS のクラスにしかないメンバーはエラー。
* **型定義にないメンバー:** どのメンバー型の型定義にもないホストのメンバーは、単独の型と同じく警告 `W0301` を出して遅延バインディングで呼ぶ。

#### 2.6.4 型の絞り込み

ローカル変数と引数は、条件式の型ガードで型が絞り込まれる。

| 型ガード | 条件が真のとき | 条件が偽のとき |
|---|---|---|
| `x instanceof T` | `T`（ユニオン型なら、`T` のインスタンスになりうるメンバー） | ユニオン型から `T` のメンバーを除いた型 |
| `typeof x == "名前"` | `TypeName` が「名前」になるメンバー（`"String"` → `string`、`"Long"` → `int` と列挙型、`"Worksheet"` → `Excel.Worksheet`） | それ以外のメンバー |

* **組み合わせ:** `!=` と `!` は真偽を入れ替える。`a && b` の `b`、`a || b` の `b` には `a` の結果が効く。`x instanceof A || x instanceof B` が真なら `A | B`。
* **効く範囲:** `if` の各分岐、条件演算子の各分岐、`while` / `for` の本体、`&&` / `||` の右辺。`if` の一方の分岐が `return` / `throw` / `break` / `continue` で抜けるときは、同じブロックの `if` より後ろの文にも効く。
  ```typescript
  if (!(s instanceof Excel.Worksheet)) return;
  s.cells(1, 1).value = 1;   // s: Excel.Worksheet
  ```
* **`any` / `object`:** `instanceof` で絞り込まれる（`activeSheet` など）。`any` は `typeof x == "String"` のような基本型の名前でも絞り込まれる。
* **絞り込まれないもの:** その範囲で代入（`=`・複合代入・`++` / `--`）される変数、モジュール変数・フィールド・`this`。絞り込まれた変数は `byref` 引数に渡せない（`E0415`。呼び出し先で書き換えられると絞り込みが成り立たなくなるため）。
* **`typeof` の名前:** `typeof` は VBA の `TypeName` と同じ名前を返す（`"String"`、`"Long"`、`"Double"`、`"Boolean"`、クラス名）。JavaScript の名前（`"string"`、`"number"` など）と比較すると一致しないため警告（`W0312`）。
* **生成する VBA:** 変数は宣言した型（`Object` / `Variant`）のままで、絞り込まれた箇所のメンバーも遅延バインディングで呼ぶ。`instanceof` は `TypeOf … Is`、`typeof` は `TypeName` になる。

#### 2.6.5 同梱の型定義のユニオン型

タイプ ライブラリが Object / Variant を返すメンバーのうち、実際に返る値の型が決まっているものは、同梱の型定義でユニオン型にしている。

| メンバー | 型 |
|---|---|
| `activeSheet`、`application.activeSheet`、`Workbook.activeSheet`、`Window.activeSheet` | `Excel.Worksheet \| Excel.Chart` |
| シートのコレクションの要素（`sheets(1)`、`for (const s of thisWorkbook.sheets)`、`Window.selectedSheets`、`sheets.add()`） | `Excel.Worksheet \| Excel.Chart` |
| `Worksheet.next`、`Worksheet.previous` | `Excel.Worksheet \| Excel.Chart` |
| `Shape.parent`、`Shapes.parent`、`ChartObject.parent`、`ChartObjects.parent`、`PageSetup.parent` | `Excel.Worksheet \| Excel.Chart` |
| `Chart.parent`（グラフ シートはブック、埋め込みグラフは `ChartObject`） | `Excel.Workbook \| Excel.ChartObject` |
| `Name.parent`（ブック単位の名前はブック、シート単位の名前はシート） | `Excel.Workbook \| Excel.Worksheet` |
| `Window.rangeFromPoint(x, y)` | `Excel.Range \| Excel.Shape` |
| `application.statusBar`（Excel が表示を管理しているときは `false`） | `string \| boolean` |
| `application.getSaveAsFilename(...)`（キャンセルは `false`） | `string \| boolean` |
| `application.getOpenFilename(...)`（キャンセルは `false`。`multiSelect` が `true` ならパスの配列） | `string \| boolean \| any[]` |

* **シートの種類:** シートのコレクションは `interface Sheets<T = Worksheet | Chart>` と宣言している（§2.4.1）。Excel 4.0 マクロシートは `Worksheet` オブジェクトとして現れる。ダイアログシート（`DialogSheet`）は型定義にないため、ダイアログシートを含むブックでは `instanceof Excel.Worksheet` の偽の分岐が `Excel.Chart` とは限らない。
* **`getOpenFilename` の配列:** Excel が返す配列は 1 始まりだが、RAS の配列はすべて 0 始まりのため（§13.3）、生成する VBA はランタイム `RAS_Rebase1` で 0 始まりに直す。`typeof` は `"Variant()"`。
* **既存のコード:** `const ws: Excel.Worksheet = activeSheet;` や `activeSheet.cells(1, 1)` はそのままコンパイルでき、警告（`W0306` / `W0305`）が出る。
* **ユニオン型にしていないもの:** 値が引数や状況でさまざまに変わるもの（`selection`、`application.inputBox`・`evaluate`・`run`・`caller`、`Range.value` など）、エラー値（`CVErr`）や `Null` を返しうるもの、引数の有無で単体とコレクションが変わるもの（`chartObjects(...)` など）は `object` / `any` のまま。

```typescript
const f = application.getOpenFilename("Excel ブック,*.xlsx", 1, "開くブック", "開く", true);
if (typeof f == "Boolean") return;          // キャンセル
if (typeof f == "Variant()") {
    for (const path of f) workbooks.open(path);
} else {
    workbooks.open(f);                     // f: string
}

for (const sh of thisWorkbook.sheets) {
    if (sh instanceof Excel.Worksheet) {
        sh.cells(1, 1).value = sh.name;
    } else {
        sh.chartTitle.text = sh.name;      // sh: Excel.Chart
    }
}
```

---

## 3. 関数定義と引数（ByVal / ByRef）

VBAのデフォルトが `ByRef` であることに起因する事故を防ぐため、RASでは**デフォルトを `ByVal`** とし、参照渡しを明示する。

* **構文:** 引数名の前に `byref`（または `byval`）修飾子を記述可能。省略時は `ByVal`。
* **可視性:** `export function` → `Public`、`export` なし → `Private`。
* **戻り値:** `return expr;` は `関数名 = expr` ＋ `Exit Function` に変換する（末尾の `return` では `Exit` を省略）。オブジェクト型の戻り値は `Set 関数名 = expr`。
* **`void` 関数:** `Sub` として出力する。

### 3.1 コード変換例

**RASコード:**
```typescript
function calculate(x: double, byref outResult: double): boolean {
    outResult = x * 2;
    return true;
}
```

**出力VBAコード:**
```vba
Private Function calculate(ByVal x As Double, ByRef outResult As Double) As Boolean
    outResult = x * 2&
    calculate = True
End Function
```

### 3.2 拡張メソッド

C# の拡張メソッドと同様に、既存の型（プリミティブ型・配列・Excel 等の COM 型・RAS のクラス）にメソッドを後付けし、`x.method()` の形で呼べるようにする。

#### 3.2.1 構文

TypeScript の `this` パラメータ構文を流用する。第1引数を `this: 型` と書いた関数が拡張メソッドになる（TS 由来のエディタ資産でそのままハイライトされる）。

```typescript
// strutil.ras
export function isBlank(this: string): boolean {
    return this.trim().length == 0;
}

export function padLeft(this: string, width: int, ch: string = " "): string {
    return this.length >= width ? this : ch.repeat(width - this.length) + this;
}

// excelutil.ras
export function lastRow(this: Excel.Worksheet, col: int = 1): int {
    return this.cells(this.rows.count, col).end(xlUp).row;
}

export function sum(this: double[]): double {
    let t = 0.0;
    for (const v of this) t += v;
    return t;
}
```

* 本体中の `this` は受け取ったオブジェクト（レシーバ）を指す。
* `byref this: T` と書くと、レシーバを参照渡しにできる（例: `function inc(byref this: int) { this++; }` → `counter.inc()`）。この場合レシーバは左辺値でなければならず、リテラルや式を渡すとエラー（§6 ルール5と同じ）。
* ジェネリクス: `function first<T>(this: T[]): T` のように型パラメータを使える（§2.4）。

#### 3.2.2 呼び出し側

```typescript
import { isBlank, padLeft } from "./strutil";
import * as xl from "./excelutil";   // 名前空間インポートでもモジュール内の拡張メソッドがすべて有効になる

if (!name.isBlank()) {
    code = id.padLeft(6, "0");
}
const last = Sheet1.lastRow();
```

**出力VBAコード:**
```vba
    If Not StrUtil.isBlank(name) Then
        code = StrUtil.padLeft(id, 6&, "0")
    End If
    last = ExcelUtil.lastRow(Sheet1, 1&)
```

* 拡張メソッドの呼び出しは、定義モジュールで修飾した通常の関数呼び出し（レシーバを第1引数）に変換する。実行時コストは通常の関数呼び出しと同じで、§8 のインライン展開・ループ内自動展開の対象にもなる（`inline function f(this: T)` も可）。
* レシーバ式は他の引数より先に、1回だけ評価する。

#### 3.2.3 解決規則

1. **型のメンバが優先:** レシーバの型（COM 型ライブラリ・RAS クラス）に同名のメソッド/プロパティがあれば、そちらを呼ぶ（C# と同じ）。拡張メソッド側には「型のメンバに隠されて呼ばれない」警告を出す。
2. **スコープ:** 同じモジュールで定義された拡張メソッドと、`import` したモジュールの `export` された拡張メソッドが候補になる。
3. **適合性:** レシーバの静的型が `this` の型と一致、または暗黙変換（§2.3 の拡大変換、配列要素型の一致、クラスの継承・インターフェイス実装）で適合するもの。
4. **最も具体的な候補を選ぶ:** 複数の候補が適合する場合、変換が少ないものを選ぶ。決まらない場合は曖昧エラー。
5. **`any` / `object` 型のレシーバ:** 静的型が不明なため拡張メソッドは解決しない（遅延バインディングのメンバ呼び出しになる）。同名の拡張メソッドがスコープにある場合は「拡張メソッドは使われない」警告を出す。
6. **`null` レシーバ:** オブジェクト型のレシーバが `Nothing` でも呼び出しは行われる（C# と同じ）。拡張メソッド内で `this == null` を判定できる。

#### 3.2.4 標準ライブラリ

* `map` / `filter` / `forEach` / `reduce` 等の配列メソッド、文字列メソッド（`trim`, `padLeft`, `startsWith` 等）、Range 用の `toArray` / `fromArray`（§9.4）は、標準ライブラリの拡張メソッドとして定義する。ユーザー定義の拡張メソッドと同じ仕組みで解決され、§4 のループ融合はこれらに対するコンパイラ組み込みの最適化として行う。
* 標準ライブラリの拡張メソッドは暗黙にインポートされる。ユーザー定義の同名・同シグネチャの拡張メソッドがスコープにある場合はユーザー定義が優先する（警告を出す）。

### 3.3 名前付き引数（v0.9で追加）

引数を `引数名: 値` の形で渡せる。省略可能な引数（既定値付き・`?` 付き）の途中を飛ばして、後ろの引数だけを指定するときに使う。

* **構文:** 引数リストの中に `name: expr` と書く。名前のない引数（位置引数）と混ぜられるが、名前付き引数の後ろに位置引数は書けない（`f(1, c: "z")` は可、`f(a: 1, 2)` は不可）。名前付き引数どうしの順序は自由。
* **使える呼び出し:** 宣言のある関数・メソッド・コンストラクタ（`new`）・拡張メソッド・Excel / VBA のメンバーとグローバル関数（`.d.ras` の引数名）・既定メンバー。引数名は宣言どおりに書く（大文字小文字を区別する。Excel / VBA の引数名も camelCase: `lookAt`、`matchCase`）。
* **使えない呼び出し:** 関数型の値（ラムダ）・`any` / `object` 型の値のメンバー（遅延バインディング）・配列や文字列などの組み込みメソッド・`new Error` / `new Map` / `new Array`。
* **拡張メソッド:** `this` 引数はレシーバで渡すため、名前では指定できない。
* **変換:** コンパイル時に宣言の順序に並べ替え、指定されなかった引数は `undefined`（§3.4 の引数の省略と同じ）として扱う。RAS の関数では既定値を埋め込み、Excel / VBA のメンバーでは引数を空欄にする（VBA の `,,`）。
* **評価順:** 引数の式は、書いた順ではなく宣言の順に評価される。副作用のある式を複数渡すときは注意する。
* **エラー:** 存在しない引数名（E0412。大文字小文字だけが違う場合は正しい名前を示す）、同じ引数の二重指定・名前付き引数の後ろの位置引数（E0413）、必須の引数の指定漏れ（E0414）、名前付き引数を使えない呼び出し（E0411）。`...` の可変長引数は名前で指定できない（E0412）。

**RASコード:**
```typescript
function label(a: int, b: int = 2, c: string = "x"): string { ... }

export function demo(ws: Excel.Worksheet) {
    const s = label(a: 123, c: "あいう");
    const r = ws.range("A:A").find("x", lookAt: xlWhole, matchCase: true);
    msgBox("完了しました", title: "集計");
}
```

**出力VBAコード:**
```vba
s = label(123&, 2&, "あいう")
Set r = ws.Range("A:A").Find("x",,, xlWhole,,, True)
MsgBox "完了しました",, "集計"
```

### 3.4 引数の省略（`undefined`）

位置で渡す引数リストの途中に `undefined` を書くと、その省略可能な引数を指定しなかったことになる（`label(1, undefined, "z")`）。扱いは名前付き引数で飛ばした引数と同じ。

---

## 4. 無名関数（ラムダ式）と配列操作

無名関数の用途は配列・コレクションの操作（`map`, `filter`, `forEach`, `reduce`, `some`, `every`, `find`, `findIndex`, `sort` 等）に特化する。

* **構文:** アロー関数 `(x) => x * 2` または `(x: double) => { return x * 2; }`。
* **キャプチャ制約（重要）:**
  * 無名関数内で参照できる外部スコープの変数は**イミュータブル（`const` または読み取り専用変数）のみ**。
  * 無名関数内での外部変数の書き換えはコンパイルエラー。

### 4.1 トランスパイル仕様（v0.2で変更）

ラムダ式は次の優先順位で変換する。

1. **ループ融合（既定）:** 配列メソッドの引数にラムダ式が**直接リテラルで**書かれている場合、クラスを生成せず、呼び出し元にループとして直接展開する。`arr.filter(...).map(...)` のようなチェーンは1本のループに融合する。
   * コールバックが受け取る添字は、JavaScript と同じく「そのメソッドを呼んだ配列の中での位置」である。融合しても変わらない。`filter` の後段（`arr.filter(f).map((x, i) => ...)` の `i`、`arr.filter(f).findIndex(g)` の結果）は、絞り込みを通った要素の中での位置になる。後段が添字を使うときだけ、`filter` ごとに通過した要素を数える。
2. **クロージャクラス（エスケープ時のみ）:** ラムダが変数に代入される・関数の引数として渡される・戻り値になる等、呼び出し位置で本体を特定できない場合のみ、従来通りクラスモジュール（`RAS_Lambda_001.cls`）を生成する。
   * キャプチャされた定数はパブリックフィールドとし、生成時にワンタイムコピーする。
   * 評価メソッド名は `Apply`。ランタイム（`RAS_Runtime.bas`）から遅延バインディングで呼び出す。`Invoke` は `IDispatch` のメンバー名と衝突し、クラスモジュールではコンパイル エラー（「このメンバー識別子は既にオブジェクト モジュールの中に存在しています」）になるため使えない。
   * リンターは「このラムダはクラスに変換されるため遅い」旨のヒント（Information）を出す。

### 4.2 コード変換例（ループ融合）

**RASコード:**
```typescript
const base = 10;
let arr = [1, 2, 3];
let result = arr.map(x => x * base);
```

**出力VBAコード:**
```vba
Public Sub Main()
    Dim base As Long
    Dim arr() As Long
    Dim result() As Long
    Dim ras_i1 As Long

    base = 10&
    ReDim arr(0 To 2)
    arr(0) = 1&: arr(1) = 2&: arr(2) = 3&
    If RAS_Len(arr) > 0 Then
        ReDim result(0 To UBound(arr))
        For ras_i1 = 0 To UBound(arr)
            result(ras_i1) = arr(ras_i1) * base
        Next ras_i1
    Else
        Erase result
    End If
End Sub
```

### 4.3 コード変換例（エスケープするラムダ）

**RASコード:**
```typescript
const base = 10;
const f = (x: double) => x * base;
let result = applyAll(arr, f);
```

**出力VBAコード (`RAS_Lambda_001.cls`):**
```vba
Public Captured_base As Long

Public Function Apply(ByVal x As Variant) As Variant
    Apply = x * Me.Captured_base
End Function
```

### 4.4 ホストのコレクション（v0.9で追加）

`worksheets` / `workbooks` / `ws.shapes` / `Excel.Range` などのホスト（COM）のコレクションでも、配列と同じ書き方でメンバーを処理できる。

```typescript
const names = worksheets.filter(ws => ws.visible == xlSheetVisible).map(ws => ws.name);
worksheets.forEach((ws, i) => { ws.cells(1, 1).value = i; });
const data = thisWorkbook.worksheets.find(ws => ws.name == "Data");   // 見つからなければ null
```

* **対象:** 既定メンバー（`(index: any): T`）と `count` を持つホスト型。`for...of` で列挙できるものと同じで、要素の型は既定メンバーの戻り値の型 `T` になる。
* **使えるメソッド:** `map` / `filter` / `forEach` / `reduce` / `some` / `every` / `find` / `findIndex` と、メンバーを配列に写す `toArray()`。コールバックの第 2 引数は 0 始まりの通し番号（`filter` の後段では、絞り込みを通ったメンバーの中での番号。§4.1）。
* **結果は配列:** `map` / `filter` / `toArray` の結果は `T[]`（コレクションではない）。以降は配列のメソッド（`length` / `join` / `sort` / `slice` など）をそのまま使える。
* **型のメンバーが優先:** ホスト型に同名のメンバーがあれば、そちらが使われる（`range.find(...)` は Excel の `Range.Find`、`range.sort(...)` は `Range.Sort`）。
* **要素の型:** `worksheets` / `book.worksheets` は `Excel.Sheets<Excel.Worksheet>`、`charts` は `Excel.Sheets<Excel.Chart>` で、コールバックの引数にシートのメンバーが補完される。グラフ シートも含む `sheets` は `Excel.Sheets`（要素は `object`）で、メンバーは遅延バインディングになる。
* **トランスパイル仕様:** §4.1 のループ融合を `For Each` に対して行う。配列は作らない（`map` / `filter` / `toArray` の結果を除く）。プロパティの連鎖で得たコレクション（`book.worksheets`）は一度だけ評価する。通し番号は、コールバックが受け取るときだけ数える。
* **列挙中の変更:** `For Each` で直接列挙するため、コールバックの中で列挙中のコレクションを変更（シートの追加・削除など）してはならない。変更する場合は `toArray()` で写してから処理する（`worksheets.toArray().forEach(ws => ws.delete())`）。

**RASコード:**
```typescript
const names = worksheets.filter(ws => ws.visible == xlSheetVisible).map(ws => ws.name);
```

**出力VBAコード:**
```vba
Dim names_() As String
Dim ras_e As Worksheet
Dim ras_n As Long

ras_n = 0&
If Worksheets.Count > 0& Then
    ReDim names_(0 To Worksheets.Count - 1&)
    For Each ras_e In Worksheets
        If ras_e.Visible = xlSheetVisible Then
            names_(ras_n) = ras_e.Name
            ras_n = ras_n + 1&
        End If
    Next ras_e
End If
If ras_n = 0& Then
    Erase names_
ElseIf ras_n <= UBound(names_) Then
    ReDim Preserve names_(0 To ras_n - 1&)
End If
```

---

## 5. エラーハンドリング

TypeScriptの `try ... catch ... finally` をベースとしつつ、VBAの `On Error GoTo` と `On Error Resume Next` を安全に表現する。

### 5.1 標準の例外捕捉（中断型）

`catch (e)` は `On Error GoTo` のラベルジャンプに変換し、エラー発生時点で `try` ブロックを中断する。

* `e` は `RAS_Error` 型（`code`, `description`, `source`, `rasFile`, `rasLine` を持つ）。catch 突入時に `Err` の内容をコピーする。
* **ハンドラの正しい脱出（罠回避）:** VBAではエラーハンドラ内から `Resume` せずに処理を続けると「ハンドラ実行中」状態が残り、以後のエラーが捕捉されない。コンパイラは catch 本体の実行後、必ず `Resume RAS_EndTry_N` でハンドラ状態を解除してから後続処理へ合流させる。
* **1プロシージャ内の複数 `try`・ネスト:** VBAはプロシージャ内で有効なハンドラが1つしかないため、プロシージャ共通のハンドラ `RAS_Handler:` を1つ生成し、現在の try 識別子 `ras_tryId` を `Select Case` で振り分ける。内側 try を抜けた時点で外側の try 識別子に戻す。
* **`throw`:** `throw new Error("msg")` は `Err.Raise vbObjectError + 513, <ソース>, "msg"`。`throw e;`（再送出）は保存した番号・説明で `Err.Raise` する。

### 5.2 エラー無視ブロック（強行突破型：`On Error Resume Next`）

`catch` ブロック内に `resume;` 文のみを記述する特殊構文。

* **リンター制約:** `catch { resume; }` 内に他のステートメントがあればエラー。
* **トランスパイル仕様:** `try` 開始時に `On Error Resume Next`、終了時に必ず `On Error GoTo 0`（外側に try がある場合は外側のハンドラを再設定）。

**RASコード:**
```typescript
try {
    let ws = worksheets("NonExistent");
    ws.activate();
} catch {
    resume;
}
```

**出力VBAコード:**
```vba
On Error Resume Next
Set ws = Worksheets("NonExistent")
ws.Activate
On Error GoTo 0
```
（`Dim ws As Object` はプロシージャ先頭に巻き上げられる。§13.1）

### 5.3 `finally` と `using`

* `finally` は、`try` の正常終了・`return`・`break`/`continue` によるブロック脱出・catch されなかったエラーの全経路で実行されるよう、脱出経路ごとに `GoTo RAS_Finally_N` と「脱出理由」変数を用いて変換する。finally 実行後、脱出理由に応じて `Exit Function` / 外側ハンドラへの再送出等を行う。
* `using x = expr;`（TS 5.2 の Explicit Resource Management 構文）を採用する。スコープ終了時に `x.dispose()` を呼ぶ `try/finally` の糖衣構文。§10 のアプリケーション状態制御で主に使う。

---

## 6. リンター（静的解析）の主要検証ルール

エディタ上でリアルタイムに検証し、`Diagnostic` を出力する。

1. **構文エラー**
2. **`const` 再代入禁止**
3. **ラムダキャプチャ違反**
4. **`resume` ブロックの純粋性検証**
5. **`byref` へのリテラル渡し禁止**
6. **（追加）大文字小文字のみ異なる識別子の衝突:** VBA は大文字小文字を区別せず、VBE が全モジュールの識別子の大小を書き換えるため、`foo` と `Foo` の共存はエラー（§13.6）。
7. **（追加）VBA 予約語・組み込み名との衝突:** 自動リネームする旨の警告（§13.6）。
8. **（追加）`inline` 関数の再帰・展開不能:** 直接・間接再帰の `inline` 関数はエラー。
9. **（追加）非修飾の `range` / `cells` / `activeSheet` 依存:** 警告（§13.4）。
10. **（追加）`.select()` / `.activate()` / `selection` の使用:** 情報（性能・安定性）。
11. **（追加）コードページで表現できない文字:** 自動エスケープした旨の情報（§13.7）。
12. **（追加）VBA の構造制限超過の予測:** プロシージャ 64KB 制限、行継続 24 個制限等（§13.8）。
13. **（追加）最適化の可視化:** インライン展開・一括転送・ループ融合が適用された／適用できなかった箇所を Hint と CodeLens で示す（理由付き）。
14. **（v0.9で追加）ホストの名前の綴り:** Excel / VBA / Office のメンバー・グローバルを VBA の綴り（`ws.Cells`、`MsgBox`）で書くと警告（`W0304`）。意味は変わらず、`ras check --fix` とクイックフィックスで camelCase に書き換えられる（§15.3.1）。
15. **（v0.9で追加）`typeof` の比較:** `typeof x == "string"` のように JavaScript の型名と比較すると警告（`W0312`）。`typeof` は VBA の `TypeName` と同じ名前を返す（§2.6.4）。

---

## 7. ツールチェイン・アーキテクチャ構成（v0.2で変更）

### 7.1 全体構成

```
┌──────────────── VS Code ────────────────┐
│ ras-vscode（薄い拡張）                    │
│  - TextMate Grammar（TS派生）            │
│  - 設定 ras.path でバイナリを指定        │
│  - `ras lsp --stdio` を起動するだけ      │
└──────────────┬──────────────────────────┘
               │ LSP (JSON-RPC over stdio)
┌──────────────▼──────────────────────────┐
│ ras（シングルバイナリ）                  │
│  parser → resolver → typecheck → HIR    │
│  → 最適化(§8,§9,§4) → VBA lowering      │
│  → emitter(.bas/.cls/ドキュメントモジュール)│
│  LSP サーバ / watch / sync クライアント  │
└──────────────┬──────────────────────────┘
               │ 共有メモリ(MMF) + COM 通知 (§14.1〜14.8)
               │ ※ブックが閉じている場合は .xlsm を直接読み書き (§14.9)
┌──────────────▼──────────────────────────┐
│ Excel  ─ RASBridge.xlam（同期用アドイン）│
│        ─ 対象ブック（.xlsm）              │
└─────────────────────────────────────────┘
```

### 7.2 実装言語

* **Rust で実装する（確定）。** 理由:
  * 依存ランタイムなしのシングルバイナリを容易に作れる（Windows x64 / arm64、CI 用に Linux / macOS も同一ソースでビルド）。
  * 同期機能で必要な Win32 API（`CreateFileMappingW`, イベントオブジェクト）と COM（ROT 経由の Excel 取得、`IDispatch` 呼び出し）を `windows` クレートで直接扱える。
  * インクリメンタル解析を前提とした LSP の応答性能。
* **パーサー:** 手書きの再帰下降パーサー（エラー回復と LSP 用の位置情報・部分木再利用を重視）。v0.1 の Chevrotain 案は採用しない。
* **採用ライブラリと参考実装:** §17 を参照。
* **プラットフォーム:** 配布対象は Windows（x64 / arm64）。ライブ同期（§14.1〜14.8）は Windows 専用。ビルド・チェック・LSP・オフライン書き込み（§14.9）は Excel にも Win32 にも依存しないため、CI 用に Linux でも動作させる。

### 7.3 CLI

| コマンド | 内容 |
|---|---|
| `ras init` | `ras.toml` と雛形を生成 |
| `ras check [--fix]` | 型チェック・リントのみ。`--fix` は機械的に直せる警告（§6 ルール 14）をソースに書き戻してから検査する |
| `ras build [--release]` | VBA ファイルを `build/` に出力。`--release` で最適化レベル既定 `-O2` |
| `ras watch` | 変更監視しながらインクリメンタルビルド＋ライブ同期 |
| `ras sync [--pull]` | 1回だけ Excel と同期（`--pull` は Excel → ローカル方向） |
| `ras pack [<book>]` | Excel を起動せず、ブックファイル内の VBA プロジェクトを直接書き換える（§14.9）。`--via-excel` で不可視の Excel 経由の書き込みに切り替え |
| `ras unpack [<book>]` | Excel を起動せず、ブックファイルから VBA ソースを `vba/` に取り出す（§14.9） |
| `ras lsp --stdio` | LSP サーバとして起動（VS Code 拡張から呼ばれる） |
| `ras setup` | 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」設定（ユーザー同意の上でレジストリ `AccessVBOM` を設定）と RASBridge のインストール |
| `ras typegen <TypeLib>` | COM タイプライブラリから RAS 型定義（`.d.ras`）を生成 |
| `ras fmt` | フォーマッタ |

### 7.4 VS Code 拡張の責務（最小）

* `ras.path` 設定（未設定時は PATH から検索）でバイナリを起動し、LSP クライアントとして接続するのみ。
* コマンドはすべて LSP の `workspace/executeCommand` に委譲: 「生成 VBA を表示」「今すぐ同期」「Excel で開く」「マクロを実行」「フォームを Excel で実行」。
* ステータスバーに同期状態（接続先ブック・未同期件数・ドリフト警告）を表示。
* IntelliSense（メンバー一覧・クイック情報・パラメーター情報等）はすべて LSP サーバ側（`ras lsp`）が提供し、拡張側は独自実装を持たない（§16）。

### 7.5 出力構成（ビルド成果物）

* `.ras` ファイルに対応する `.bas` モジュール、`class` 宣言に対応する `.cls`
* `ThisWorkbook` / 各シートモジュールに注入するイベントハンドラコード（§12）
* エスケープしたラムダの `RAS_Lambda_XXX.cls` 群
* ランタイムライブラリ `RAS_Runtime.bas`（実際に使用された関数のみ出力する tree shaking を行う）
* ソースマップ `build/*.rasmap`

---

## 8. 最適化: インライン展開とループ内自動展開

VBA は関数呼び出しのオーバーヘッドが大きい（特に引数が `Variant` の場合やモジュールを跨ぐ場合）。RAS は以下の2つの仕組みで呼び出しを削減する。

### 8.1 `inline` キーワード（明示的インライン）

```typescript
inline function clamp(v: double, lo: double, hi: double): double {
    return v < lo ? lo : (v > hi ? hi : v);
}
```

* `inline` 関数の呼び出しはすべて呼び出し元に展開される。最適化レベルに関係なく必ず展開する（`-O0` を除く。`-O0` ではデバッグ容易性を優先して通常呼び出しにする）。
* `export inline function` の場合、他モジュールやマクロ一覧から呼べるよう、通常の関数定義も併せて出力する。非 export の inline 関数は定義自体を出力しない（全呼び出しが展開済みのため）。
* 直接・間接再帰はコンパイルエラー。
* `noinline function` で自動展開（§8.2）を抑止できる。

### 8.2 ループ内の自動展開

`-O1` 以上で、ループ本体（ネストしたループを含む）内の呼び出しのうち、以下をすべて満たすものを自動展開する。

* 呼び出し先が同一プロジェクト内の RAS 関数で、ソースが利用可能
* 再帰していない、`noinline` でない
* 本体のコスト（HIR ノード数）が閾値以下（`-O1`: 40、`-O2`: 120、`ras.toml` で変更可）
* 呼び出し先に `On Error Resume Next` 型の try、`Static` 相当の状態、`using` を含まない（含む場合は展開結果のエラー状態が呼び出し元に漏れるため）
* 呼び出し先が別モジュールの `Private` なモジュール変数／関数を参照していない（展開先から見えないため）
* 展開によって呼び出し元プロシージャが VBA の 64KB 制限に近づかない（§13.8。超える場合は展開を打ち切り、Hint で通知）

ループ外の呼び出しは、`-O2` でかつ本体が非常に小さい（コスト 10 以下）場合のみ展開する。

### 8.3 展開の意味論（保証事項）

展開後もソース上の評価順序・副作用・値渡し/参照渡しの意味を厳密に保つ。

* **値渡し引数:** 実引数を一時変数に一度だけ評価して代入する。ただし実引数が「副作用のない単純変数またはリテラル」で、かつ仮引数が本体で再代入されない場合は直接置換して一時変数を省略する。
* **`byref` 引数:** 実引数の左辺値で置換する。左辺値の添字に式が含まれる場合（`a(f(i))` など）は添字を先に一時変数へ評価して固定する。
* **ローカル変数:** `ras_inl{N}_{名前}` にリネームし、呼び出し元プロシージャの先頭に宣言を巻き上げ、展開位置で明示的に初期化する（§13.1）。
* **早期 `return`:** 本体の末尾以外の `return` は、結果変数への代入＋展開ごとに一意なラベル `RAS_InlEnd_{N}` への `GoTo` に変換する。
* **式中の呼び出し:** 展開対象の呼び出しが式の中にある場合、コンパイラはその文を A正規形（ANF）に変換し、短絡評価（`&&`, `||`）・三項演算子・ループ条件の評価タイミングを保ったまま文の列に展開する。ループ条件中の呼び出しは `Do ... If Not cond Then Exit Do ... Loop` 形に変換する。

### 8.4 コード変換例

**RASコード:**
```typescript
for (let i = 0; i < n; i++) {
    total += clamp(values[i], 0, 100);
}
```

**出力VBAコード:**
```vba
    For i = 0 To n - 1
        ' @inline clamp (util.ras:1)
        If values(i) < 0 Then
            ras_inl1_ret = 0
        ElseIf values(i) > 100 Then
            ras_inl1_ret = 100
        Else
            ras_inl1_ret = values(i)
        End If
        total = total + ras_inl1_ret
    Next i
```

### 8.5 その他のループ最適化

* **`for` の `For...Next` 化:** `for (let i = a; i < b; i++)` は、`i` と `b` が本体で変更されず、ステップが定数の場合に `For i = a To b - 1` へ変換する（VBA の `For` は上限を1回だけ評価するため、JS の毎回評価と意味が一致する条件を検証する）。満たさない場合は `Do While` に変換する。
* **ループ不変式の巻き上げ:** `arr.length`, `obj.count`, 同一オブジェクトのプロパティチェーン（`ws.range("A1").font` 等）でループ内で不変と判定できるものはループ前に一時変数へ退避する。COM オブジェクトのプロパティは「ループ内でそのオブジェクトに書き込み・メソッド呼び出し・不明な関数呼び出しがない」場合のみ不変とみなす。
* **`With` ブロック化:** 同一オブジェクト式への連続したメンバアクセスは `With` に変換する。
* **文字列連結:** ループ内での `s += ...` は、ランタイムの文字列バッファ（`Space$` で確保した領域に `Mid$` で書き込む方式）に自動変換する（`-O2`）。
* **`$` 付き関数:** `Left`, `Mid`, `Trim` 等は `String` 型を返す `$` 版（`Left$` 等）を出力する。

---

## 9. Range ⇔ 2次元配列の自動一括変換

セル1つずつの読み書きは1回ごとに COM 呼び出しが発生し極めて遅い。RAS はループ内のセルアクセスパターンを検出し、**Range と2次元配列の一括転送**に自動変換する。

### 9.1 書き込み（ループ → 配列 → Range 一括代入）

**RASコード:**
```typescript
const ws = Sheet1;
for (let r = 2; r <= last; r++) {
    for (let c = 1; c <= 5; c++) {
        ws.cells(r, c).value = r * c;
    }
}
```

**出力VBAコード:**
```vba
    If last >= 2 Then
        ReDim ras_buf1(0 To last - 2, 0 To 4)
        For r = 2 To last
            For c = 1 To 5
                ras_buf1(r - 2, c - 1) = r * c
            Next c
        Next r
        ws.Range(ws.Cells(2, 1), ws.Cells(last, 5)).Value = ras_buf1
    End If
```

### 9.2 読み込み（Range 一括取得 → 配列参照）

**RASコード:**
```typescript
let sum = 0;
for (let r = 2; r <= last; r++) {
    sum += ws.cells(r, 1).value;
}
```

**出力VBAコード:**
```vba
    If last >= 2 Then
        ras_src1 = RAS_Read2D(ws.Range(ws.Cells(2, 1), ws.Cells(last, 1)))
        For r = 2 To last
            sum = sum + ras_src1(r - 2, 0)
        Next r
    End If
```

`RAS_Read2D` は「単一セルの Range の `.Value` は配列ではなくスカラーを返す」VBA の罠を吸収し、常に **0 始まり**の2次元配列を返す。

### 9.2.1 添字規則（0 始まり）

* RAS の配列はすべて 0 始まり（§13.3）であり、Range 由来の2次元配列も例外としない。自動変換（§9.1, §9.2）・明示 API（§9.4）とも、RAS コードから見える添字は `[行 - 先頭行][列 - 先頭列]` の 0 始まりとする。
* **ゼロコピーでの再基底化:** `Range.Value` が返す配列は 1 始まりである。要素をコピーし直すと大きな範囲で無視できないコストになるため、`RAS_Read2D` は配列本体をコピーせず、`SAFEARRAY` 記述子の `rgsabound[].lLbound` を直接 0 に書き換える（`RtlMoveMemory` 使用。`Variant` 内の `SAFEARRAY*` の取得、記述子内オフセットの 32bit / 64bit 差は §11 の規則で両方出力する）。
  * 書き換えは、`cLocks = 0`・`cDims = 2`・`fFeatures` に `FADF_STATIC` / `FADF_FIXEDSIZE` が立っていない場合に限る。条件を満たさない場合は要素コピーにフォールバックする。
  * 書き込み方向は、`Range.Value` への代入が下限 0 の配列をそのまま受け付けるため、変換不要。
* **単一行・単一列:** 1 列の Range も2次元配列（`n × 1`）として扱う。1 次元配列が欲しい場合は `range.toColumn()` / `range.toRow()` を使う（こちらはコピーを伴う）。

### 9.3 変換の適用条件

コンパイラは以下をすべて満たす場合のみ自動変換する（満たさない場合は変換せず、理由を Hint で表示）。

1. **対象シート式がループ内で不変**（`const` 変数、コード名、またはループ外で確定したオブジェクト）。
2. **添字が誘導変数の一次式**（`r + k`, `c - k` の形。`k` はループ不変）で、ループ範囲から矩形領域が静的に（ループ開始前に）確定する。
3. **アクセスするプロパティが `value` / `value2` / `formula` / `formulaR1C1` のいずれか**で、変換前後で同じプロパティを使う（`.Value` を `.Value2` に勝手に変えない。日付・通貨の値が変わるため）。
4. **ループ内で同じシートに対する他の操作がない**: 書式変更・`Select`・行列の挿入削除・他のセルアクセス・シートに触れうる不明な関数呼び出し（展開できなかった呼び出し、`Application.Run`、COM メソッド呼び出し）がない。
5. **書き込みの場合、矩形内の全セルに無条件で書き込む**（条件付き書き込み・`continue`・`break` がない）。部分的な書き込みを一括化すると、書き込まないセルの数式が値で上書きされるため。
6. **書き込み領域を同じループ内で読み戻していない**（読み戻しは配列参照に置換できる場合のみ許可）。
7. **`break` / `return` / 例外による途中脱出がない**（書き込みの場合。読み込みは途中脱出可）。

### 9.4 明示的な制御

* `@bulk` をループに付けると、条件 4〜6 の一部（副作用解析で保守的に不可と判断されたもの）をユーザーの責任で無視して変換する。条件 5 を満たさない場合は、事前に `.Formula` で既存内容を読み込んでから書き戻す方式を使う。
* `@nobulk` で変換を抑止する。
* 明示 API も提供する:
  * `range.toArray(): any[][]` — 常に 0 始まりの2次元配列（§9.2.1）
  * `range.fromArray(arr)` — 配列サイズに合わせて Range をリサイズして一括代入（配列の下限は問わない）
  * `sheet.usedTable(): any[][]` — 使用範囲を一括取得

---

## 10. アプリケーション状態の制御（1行 API）

画面更新・イベント・自動計算等の停止/再開を1行で行えるランタイム API を提供する。

### 10.1 API 一覧

| RAS | 内容 |
|---|---|
| `using _ = fastMode();` | 画面更新・イベント・自動計算・警告表示・ステータスバーの表示を停止し、スコープを抜けると（エラー時も）元の状態に戻す |
| `using _ = fastMode({ calc: false });` | 個別に除外指定可（`screen`, `events`, `calc`, `alerts`, `statusBar`, `cursor`）。`interactive: false` はユーザー入力の遮断を追加する（既定は遮断しない。§10.4.5） |
| `@fast export function f() {...}` | 関数全体に `fastMode()` を適用するデコレータ |
| `suspendScreen()` / `resumeScreen()` | `ScreenUpdating` |
| `suspendEvents()` / `resumeEvents()` | `EnableEvents` |
| `suspendCalc()` / `resumeCalc()` | `Calculation`（`xlCalculationManual` ⇔ 元の値） |
| `suspendAlerts()` / `resumeAlerts()` | `DisplayAlerts` |
| `suspendAll()` / `resumeAll()` | 上記すべて |
| `status("処理中 50%")` | `StatusBar` 表示（`fastMode` 終了時に `False` に戻す） |

### 10.2 ランタイムの振る舞い

* **ネスト対応:** 各項目ごとに深さカウンタを持つ。深さ 0→1 のときだけ元の状態を保存して停止し、1→0 のときだけ保存した状態に戻す。`fastMode` を使う関数から `fastMode` を使う関数を呼んでも、内側の終了時に画面更新が再開されない。
* **元の状態を保存して戻す:** 常に `True` / `xlCalculationAutomatic` に戻すのではなく、停止前の値に戻す。
* **計算モードの罠回避:** 開いているブックがない状態で `Application.Calculation` を設定するとエラーになるため、ランタイムがガードする。
* **エラー時の復元保証:** `using` は `finally` 経由で必ず復元する。加えて、マクロ入口（§12.3）の最上位ハンドラが全カウンタを 0 に戻して状態を復元するため、未捕捉エラーで `EnableEvents = False` のまま残る事故が起きない。

### 10.3 コード変換例

**RASコード:**
```typescript
@fast
export function rebuildReport() {
    // ...
}
```

**出力VBAコード:**
```vba
Public Sub rebuildReport()
    RAS_EnsureInit
    On Error GoTo RAS_Handler
    RAS_SuspendAll
    ' ... 本体 ...
RAS_Finally_0:
    RAS_ResumeAll
    Exit Sub
RAS_Handler:
    RAS_ResetAppState
    RAS_ReportError "Report.rebuildReport", Err
End Sub
```

---

### 10.4 `DoEvents` の自動挿入

長時間ループの間に Excel が「応答なし」になるのを防ぎ、Esc キーで中断できるようにするため、コンパイラが**止まりそうな箇所に `DoEvents` を間引いて自動挿入**する。ユーザーは `DoEvents` を書かなくてよい。

#### 10.4.1 挿入箇所

ループの末尾（次の繰り返しへ戻る直前）に挿入する。対象は次のいずれかに当たるループ。

* `while` / `do-while` / 無限ループ（`for (;;)`）
* 回数が実行時まで決まらない `for` / `for-of`
* 回数がコンパイル時に分かり、かつ閾値以上（既定 10,000 回）の `for`
* 本体に COM 呼び出し、ファイル I/O、またはループを含む関数呼び出しを含むループ

ネストしたループの扱い:

* 原則として最も内側の対象ループに入れる。
* ただし最も内側のループの回数が定数で小さい場合（`for (let c = 0; c < 5; c++)` など）は、1つ外側のループに入れる。
* 1つのループの入れ子に入れるのは1か所だけとする。

ループ以外では、標準ライブラリの `sleep(ms)` / `waitUntil(...)` が待機を短く分割し、その間で `DoEvents` を呼ぶ。

#### 10.4.2 間引き

`DoEvents` は1回の呼び出しが重いため、2段階で間引く。

1. **カウンタ判定（ループ内に展開する）:** 1回の加算とビット演算だけの判定。`RAS_Yield` の呼び出しは N 回に1回に抑える。
2. **時間・入力判定（`RAS_Yield` 内）:**
   * 前回の `DoEvents` から一定時間（既定 100ms）が経ったとき、または `GetInputState()` が保留中の入力（マウス・キー）を報告したときだけ `DoEvents` を呼ぶ。
   * `RAS_Yield` は前回の呼び出しからの経過時間を見て N を自動調整する。概ね 10ms ごとに呼ばれるよう、N は 2 のべき乗の範囲で増減させる。

**出力VBAコード（ループ側）:**
```vba
    Do While ras_cond1
        ' ... 本体 ...
        ras_y1 = ras_y1 + 1&
        If (ras_y1 And ras_yieldMask) = 0& Then RAS_Yield
    Loop
```

**出力VBAコード（ランタイム側、概略）:**
```vba
Public Sub RAS_Yield()
    Dim now_ As Double
    now_ = RAS_TickMs()                ' GetTickCount の 49.7 日の折り返しを補正した値
    RAS_AdjustYieldMask now_           ' 前回呼び出しからの間隔で ras_yieldMask を増減
    If GetInputState() <> 0 Or now_ - ras_lastYield >= RAS_YIELD_MS Then
        ras_lastYield = now_
        DoEvents
    End If
End Sub
```

* `GetTickCount` は `Long` で返るため、そのまま引き算すると折り返し時にオーバーフロー（エラー 6）になる。ランタイムは `double` に変換し、折り返しを補正して扱う。
* `GetInputState` / `GetTickCount` の `Declare` は §11 の規則で 32bit / 64bit 両方を出力する。

#### 10.4.3 挿入しない箇所

* **ワークシート関数（UDF、`@udf`）と、そこから呼ばれうる関数:** セルの計算中に `DoEvents` を呼ぶと、再計算の不整合や異常終了を招くため。呼び出しグラフで判定する。
* **イベントハンドラ（`@on`）と、そこから呼ばれうる関数:** イベント処理中の再入を防ぐため。`@doevents` を付ければ挿入する。
* **`@init`、クラスの `Class_Terminate`**
* **§9 の一括転送に変換したループ:** 変換後は配列操作だけになり十分速い場合が多いため。回数が閾値以上なら §10.4.1 の規則で挿入する。

#### 10.4.4 明示的な制御

* `@nodoevents`: 関数・ループに付けて挿入を抑止する。
* `@doevents`: 関数・ループに付けて挿入を強制する。
* `yieldNow()`: その場で `RAS_Yield` を呼ぶ。
* `ras.toml` の `[doevents]` で全体の有効・無効、間隔、回数閾値を設定する（§15.2）。
* LSP は、挿入した箇所を CodeLens で表示する。

#### 10.4.5 `DoEvents` 中の再入と中断

`DoEvents` の間はユーザー操作やイベントが処理されるため、ランタイムは次の対策を自動で行う。

* **マクロの再入防止:** 実行中フラグ `RAS_Busy` を持つ。実行中にボタンやリボンから別のマクロ入口（§12.4）が呼ばれた場合は、実行せずステータスバーに「処理中です」と表示する。`@reentrant` を付けたマクロは除く。
* **イベント:** `fastMode` の `events` で停止していればイベントは発生しない。停止していない場合、`DoEvents` 中に発生したイベントのハンドラは通常どおり実行する。ただし、自身が実行中のハンドラへの再入は §12.2 の再入防止で無視する。
* **Esc キーでの中断:** マクロ入口で `Application.EnableCancelKey = xlErrorHandler` を設定する。`DoEvents` 中に Esc が押されるとエラー 18 として最上位ハンドラに届き、アプリケーション状態を復元して「ユーザーにより中断されました」と表示して終了する。`using` / `finally` は中断時も実行する。`@nocancel` で中断を無効にできる。
* **セル編集モードとの競合:** `DoEvents` 中にユーザーがセル編集を始めると、その後のシート操作が失敗しうる。`fastMode({ interactive: false })`（`Application.Interactive = False`）を指定すると、ユーザー入力を受け付けず、再描画と「応答なし」の回避だけを行う。

## 11. `declare`（外部 DLL 関数）の 32bit / 64bit 両対応

### 11.1 構文

TS の `declare module` 構文を転用する。

```typescript
declare module "user32" {
    function GetForegroundWindow(): ptr;
    @alias("GetWindowTextW")
    function GetWindowText(hWnd: ptr, lpString: wstr, nMaxCount: int): int;
}
declare module "kernel32" {
    function Sleep(dwMilliseconds: int): void;
    function GetTickCount64(): long;  // 64bit 専用 → §11.3
}
```

### 11.2 宣言用の型

| RAS 型 | VBA7 側 | VBA6 側 | 備考 |
|---|---|---|---|
| `ptr` | `ByVal x As LongPtr` | `ByVal x As Long` | ハンドル・ポインタ |
| `int` | `ByVal x As Long` | 同左 | |
| `short` | `ByVal x As Integer` | 同左 | |
| `byte` | `ByVal x As Byte` | 同左 | |
| `long` | `ByVal x As LongLong` | 使用不可 | |
| `float` | `ByVal x As Single` | 同左 | |
| `double` | `ByVal x As Double` | 同左 | |
| `wstr` | `ByVal x As LongPtr` | `ByVal x As Long` | 呼び出し側で自動的に `StrPtr(s)` を渡す（W 版 API 用） |
| `astr` | `ByVal x As String` | 同左 | ANSI 文字列（A 版 API 用） |
| `byref T` | `ByRef x As T` | 同左 | 出力引数・構造体 |

### 11.3 出力

すべての `declare` について、コンパイラは VBA7（PtrSafe・`LongPtr`）用と VBA6 用の両方を条件付きコンパイルで出力する。

```vba
#If VBA7 Then
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal nMaxCount As Long) As Long
#Else
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
#End If
```

* `long` を含む宣言は `#If Win64 Then` 内にのみ出力し、32bit 環境で呼ばれた場合は実行時エラー（または `@fallback` で指定した代替実装）にする。リンターは `if (Platform.is64bit)` ガードの外で呼んでいれば警告する。
* ユーザーコード中の `ptr` 型変数の `Dim` も、VBA6 互換を有効にしている場合は `#If VBA7` で `LongPtr` / `Long` を出し分ける。
* `ras.toml` の `target.vba6 = false` で VBA6 用の出力を省略できる（既定は `true`）。
* 構造体（`declare struct`）は 64bit でのアラインメント差を考慮し、ポインタサイズに依存する `LenB` 等の計算をコンパイル時に両方用意する。
* 頻出 API（`Sleep`, `GetTickCount`, `CopyMemory` 等）は標準ライブラリ `import { Sleep } from "win32";` として提供し、使用されたものだけを出力する。

---

## 12. ボイラープレートの自動生成

### 12.1 モジュール共通

全モジュールに以下を自動付与する。ユーザーは記述しない。

* `Attribute VB_Name = "..."`（ファイル名から生成）
* `Option Explicit`
* `Option Compare Binary`（明示。比較の挙動をモジュール依存にしない）
* `Option Private Module`（`@macro` も `export` も持たないモジュール。マクロ一覧の汚染防止）
* クラスの `VB_PredeclaredId` 等の属性（`@predeclared` 等のデコレータから生成）

### 12.2 イベントハンドラ

ドキュメントモジュール（`ThisWorkbook`, シートモジュール）はファイルとしてインポートできないため、コンパイラがイベントハンドラコードを生成し、同期機能（§14）がコードモジュールに直接書き込む。

```typescript
@on("Workbook.Open")
export function setup() { ... }

@on("Sheet1.Change")
function onSheet1Change(target: Excel.Range) { ... }

@on("Workbook.SheetChange")
function onAnySheetChange(sh: object, target: Excel.Range) { ... }
```

* 同じイベントに複数のハンドラを登録可能（宣言順に呼ぶ）。
* イベントハンドラは自動的に再入防止（ハンドラ実行中に自身が発火させたイベントを無視する）を行う。`@on("...", { reentrant: true })` で解除。
* シートはコード名で指定する。

### 12.3 初期化の保証（`@init`）

`@init` 関数は「VBA プロジェクトの状態ごとに、確実に1回だけ」実行される。

`Workbook_Open` だけに初期化を書くと、次のケースで初期化されないか失敗する:

* マクロ無効状態・保護ビューで開かれ、後から「コンテンツの有効化」「編集を有効にする」を押した場合の環境差・タイミング差（ウィンドウ未準備で `ActiveWindow` 等が使えない等）
* 他のマクロのエラーで `Application.EnableEvents = False` が残っている
* Shift キーを押しながら開いた、別ブックの VBA から `Workbooks.Open` された（`Auto_Open` は呼ばれない）
* VBE でのコード編集・未捕捉エラー・`End` による**状態リセット**でグローバル変数（リボンの `IRibbonUI` 参照等）が消えた

RAS は以下のすべての経路から `RAS_EnsureInit` を呼び、初期化済みフラグで1回に制限する。

1. `Workbook_Open`（即時実行せず `Application.OnTime Now, "RAS_DeferredInit"` で遅延実行し、ウィンドウ準備前の失敗を避ける）
2. `Auto_Open`
3. `Workbook_Activate` / `Workbook_WindowActivate`
4. すべてのマクロ入口（`@macro` / `export` された引数なし Sub、ボタン・リボンのコールバック）の先頭
5. リボンの `onLoad` コールバック

初期化処理の中で `EnableEvents` が `False` で残っていれば警告ログを出して `True` に戻す（`ras.toml` で無効化可能）。

### 12.4 マクロ入口の自動ラップ

ユーザーから直接呼ばれうる関数（マクロ入口）には自動的に以下を付与する。

* `RAS_EnsureInit` 呼び出し
* 再入防止（`RAS_Busy`、§10.4.5）と `Application.EnableCancelKey = xlErrorHandler` の設定
* 最上位エラーハンドラ: アプリケーション状態の復元（§10.2）→ RAS ソース位置付きのエラーダイアログ表示（`@silent` で抑止、`@onError(handler)` で差し替え）

### 12.5 リボン参照の復元

`@ribbon` 指定時、`IRibbonUI` の参照を `ObjPtr` で隠し名前定義（`Workbook.Names` の非表示名）にも保存し、状態リセット後でも `RAS.ribbon.invalidate()` が動作するよう自動復元する（`CopyMemory` による復元コードは §11 の 32/64bit 両対応で出力）。

---

## 13. VBA の罠の自動回避

以下はユーザーが意識しなくてもコンパイラ・ランタイムが自動で回避する。

### 13.1 宣言・初期化

* **ループ内の `Dim` は再初期化されない:** VBA の `Dim` はプロシージャ単位で1回しか初期化されない。RAS は全宣言をプロシージャ先頭に巻き上げ、RAS 上の宣言位置で明示的に初期化値（`0`, `""`, `Nothing`, `Erase` 等）を代入する。
* **`Dim a, b As Long` で `a` が `Variant` になる:** 1行1変数で出力する。
* **`Dim x As New Foo` の自動インスタンス化:** 使用しない。`new Foo()` は常に `Set x = New Foo` に変換する。
* **ブロックスコープ:** RAS のブロックスコープ変数で名前が重複する場合は、出力時にリネームする（`x`, `x_2`）。

### 13.2 演算・評価

* **`And` / `Or` は短絡評価しない:** `&&` / `||` はネストした `If` と一時変数に変換する。
* **`IIf` は両辺を評価する:** 三項演算子は `If` 文に変換し `IIf` は使わない。
* **整数リテラルの型:** VBA は小さい整数リテラルを `Integer` とみなすため `300 * 200` が溢れる。数値リテラルには常に型サフィックスを付与する（§2.2）。`short` 同士の演算は `int` に昇格させる（§2.3）。
* **`True = -1`:** `b as int` 等の boolean の数値化は `-CLng(b)` に変換して `1` / `0` にする（C# の `b ? 1 : 0` と同じ値）。
* **`Mod` は実数を丸めてから計算する:** `%` はオペランドが整数型なら `Mod`、それ以外はランタイム `RAS_FMod`（§2.3）。
* **`Round` は銀行丸め:** `Math.round` は 0 から遠い方向への四捨五入（`MidpointRounding.AwayFromZero` 相当）のランタイム関数に変換する。`Math.bankersRound` で VBA 準拠（偶数丸め）も選べる。
* **`/` と `\`:** 整数同士の `/` は `\`、それ以外は `/` に変換する（§2.3）。VBA の `\` は実数オペランドを銀行丸めしてから割るため、実数を含む整数除算は出力しない。
* **`==` / `===`:** オブジェクトなら `Is`、それ以外は `=`。`null` / `undefined` との比較はオブジェクト型なら `Is Nothing`、`any` なら `RAS_IsNullish`（`IsEmpty`, `IsNull`, `Is Nothing` を判定）。
* **既定メンバの暗黙呼び出し:** `rng` を値として使う場合も `.Value` を明示出力する。

### 13.3 文字列・配列・コレクション

* **文字列は RAS で 0 始まり:** `s.substring(i, j)`, `s[i]`, `s.indexOf(x)` は `Mid$`, `InStr` の 1 始まりに自動補正する。
* **配列は常に 0 始まり:** `Option Base` に依存しないよう、配列生成は常に下限を明示（`ReDim a(0 To n - 1)`）する。
* **空配列に `UBound` でエラー:** `arr.length` はランタイム `RAS_Len`（未割り当て・`Erase` 済みでも 0 を返す）に変換する。`For` 変換時は長さ 0 のガードを付ける。
* **`ReDim Preserve` の O(n²):** ループ内の `push` は容量倍増方式に変換する（非公開の容量変数を持ち、ループ後に実サイズへ切り詰める）。
* **`Collection` は 1 始まり:** RAS の `Collection` ラッパは 0 始まりで添字を補正する。
* **`Scripting.Dictionary`:** 遅延バインディング（`CreateObject`）で生成し、参照設定に依存しない。`Map<K, V>` 型として提供する。
* **`WorksheetFunction.Match` 等が見つからないとき実行時エラー:** RAS の `Excel.match()` 等は `Application.Match` を用い、エラー値を `null` に変換して返す。
* **`Range.Find` の設定持ち越し:** `find()` ラッパはすべての引数（`LookAt`, `LookIn`, `MatchCase` 等）を毎回明示する。

### 13.4 シート・Range の参照

* **非修飾の `cells` / `range`:** VBA の `ws.Range(Cells(1, 1), Cells(2, 2))` の内側 `Cells` は ActiveSheet を指してエラーになる。RAS では `ws.range(...)` の引数内の非修飾 `cells` / `range` を `ws` で自動修飾する。それ以外の非修飾使用は警告。
* **シートの参照:** 名前文字列（`worksheets("Sheet1")`）よりコード名を推奨し、コード名は型付きで補完される。`worksheets(...)` の結果は `Excel.Worksheet` として扱う（タイプ ライブラリ上は `Object`。§4.4）。

### 13.5 制御構文

* **`continue` が無い:** `GoTo RAS_Continue_N` に変換する。
* **多重ループの `break label`:** `Exit For` は1段しか抜けられないため `GoTo` に変換する。
* **`switch` のフォールスルー:** `Select Case` に変換し、フォールスルーのある `case` は `GoTo` で表現する。`case` の比較は `===` 意味論。

### 13.6 識別子

* **大文字小文字の非区別:** VBA は大小を区別せず、VBE が宣言時の大小で全モジュールを書き換える（`Value` を変数名にすると `.value` が全部書き換わる等）。RAS は大小のみ異なる識別子をエラーにする。また VBA 組み込みメンバ名（`Value`, `Name`, `Count`, `Item` 等）や同梱の型定義のグローバル（`Range`, `Workbooks`, `Left` 等）とローカル変数名が衝突する場合は自動リネームする（`value` → `value_`、`const range = ws.range("A1")` → `range_`）。
* **ホストの名前の綴り:** RAS のソースでは camelCase（`ws.cells`）、生成する VBA ではタイプ ライブラリの綴り（`ws.Cells`）を使う（§15.3.1）。
* **予約語・組み込み関数名との衝突:** `Date`, `Name`, `Type`, `Error`, `Len` 等と同名の識別子は自動リネームする。
* **公開名の衝突:** 複数モジュールで同名の `Public` 関数がある場合、呼び出しを常に `モジュール名.関数名` で修飾して出力する。
* **文字列で指定される名前:** `Application.Run "name"`, `OnTime`, リボンコールバック、`@on` で生成されるハンドラ名は、存在確認を行い、リネーム・インライン除去の対象外にする。

### 13.7 文字コード

* VBE のコードモジュールはシステムのコードページ（日本語環境では CP932）で保存される。コードページで表せない文字は `?` に化ける。RAS は文字列リテラル中の表現不能文字を `ChrW(&H...)` の連結に、コメント中の表現不能文字を `\u{...}` 表記に自動変換する。
* 出力のターゲットコードページは `ras.toml` の `target.codepage`（既定 `932`）で指定する。
* `build/` の `.bas` / `.cls` は target のコードページ・CRLF で出力する。VBE の「ファイルのインポート」はファイルを**実行中の Windows の ANSI コードページ**（`GetACP()`）で読み、`VBA プロジェクトのコードページ`を見ないため、`target.codepage` と異なるロケールの PC で手動インポートすると文字化けする。また改行が LF のみの `.cls` は、VBE がクラスではなく標準モジュールとして取り込む。手動インポートではなく `ras pack` を使う（オフライン書き込みはプロジェクトのコードページを明示するので、この問題は起きない）。

### 13.8 VBA の構造上の制限

* **1プロシージャ 64KB 制限:** 出力サイズを見積もり、超える場合はインライン展開を打ち切る。それでも超える場合はエラー。
* **行継続 `_` の 24 個制限:** 長い式は一時変数に分割する。
* **1行 1023 文字制限:** 長い文字列リテラルは連結に分割する。
* **ローカル変数の上限:** 巻き上げと展開で増えた一時変数は型ごとに再利用する（生存区間が重ならないもの）。

---

## 14. Excel との同期（ライブ同期／オフライン書き込み）

### 14.1 目的と方針

* Excel が起動中でも、`.ras` の保存から数百ミリ秒以内に対象ブックの VBA プロジェクトへ反映する。
* 従来の「ファイルにエクスポート → `VBComponents.Import`」方式はディスク I/O と1モジュールごとの COM 呼び出しで遅いため、**共有メモリ（メモリマップドファイル）で一括転送し、Excel 内のアドインが VBA 側から直接コードモジュールを書き換える**方式を採る。

### 14.2 構成要素

> **実装メモ（現行版）:** 現行の実装は RASBridge と共有メモリを使わず、ras（`ras sync` / `ras watch` / LSP サーバの同期スレッド）が COM 経由で VBE のオブジェクト モデル（`VBComponents`、`CodeModule`）を直接操作する。接続（§14.3 の 1）、適用規則（§14.6）、ドリフトの扱い（§14.7）は本節以降の記述どおりで、「Bridge」を「ras の同期スレッド」と読み替える。共有メモリによる転送（§14.4）と Bridge 側のポーリング（§14.5）は、同じ規則の高速化として後から追加する。

* **`ras watch`（または LSP サーバ内の同期スレッド）:** 変更を検出してインクリメンタルビルドし、差分を共有メモリに書き込む。
* **RASBridge.xlam:** `ras setup` でインストールする Excel アドイン（RAS 自身で記述・ビルドする）。共有メモリを読み、`VBProject` を書き換え、結果を共有メモリに書き戻す。
* **前提:** 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が有効であること。`ras setup` がユーザー同意の上で設定する。

### 14.3 接続

1. `ras.toml` の `workbook` に指定されたブック（フルパス）を、COM の ROT（Running Object Table）から実行中の Excel インスタンスを列挙して探す。複数インスタンスにも対応する。
2. 見つかった場合、共有メモリ `Local\RAS.<workspaceId>`（`workspaceId` はブックのフルパスとランダムなセッショントークンから生成）を作成し、COM 経由で `Application.Run "RASBridge.xlam!RAS_Connect", <名前>, <トークン>` を呼ぶ。
3. ブックが開かれていない場合: `ras watch` は待機し、ブックが開かれた時点で接続する（RASBridge の `Application.WorkbookOpen` イベントから通知）。`ras pack` はファイルへ直接書き込む（§14.9）。

### 14.4 共有メモリのレイアウト

```
Offset  Size  内容
0       4     magic "RAS1"
4       4     protocol version
8       8     書き込み側シーケンス番号（ras が増やす）
16      8     確認済みシーケンス番号（Bridge が増やす）
24      4     状態（0=空, 1=送信済み, 2=適用中, 3=完了, 4=失敗）
28      4     ペイロード長（バイト）
32      ...   ペイロード
```

* 既定サイズ 64MB（`ras.toml` で変更可）。
* ペイロードは UTF-16LE。VBA 側は `OpenFileMappingW` / `MapViewOfFile` / `RtlMoveMemory` で、`Space$(n)` で確保した `String` の `StrPtr` 領域へ直接コピーするため、文字列変換コストがほぼゼロになる。
* ペイロードはモジュール単位のレコード列: `{操作(作成/更新/削除), モジュール種別, モジュール名, 内容ハッシュ, 本文}`。

### 14.5 通知方式

* **プッシュ（主）:** ras が書き込み後、COM で `Application.Run "RASBridge.xlam!RAS_Pull"` を呼ぶ。Excel がセル編集中・モーダルダイアログ表示中で呼び出しが拒否された場合（`RPC_E_CALL_REJECTED` 等）は指数バックオフで再試行する。
* **ポーリング（従）:** Bridge は `Application.OnTime` で1秒ごとにシーケンス番号だけを確認し、プッシュが失敗し続けた場合も取りこぼさない。
* COM 経由で送るのは短い通知のみで、コード本体は必ず共有メモリ経由で渡す。
* `Application.Run` に渡すマクロ名は、必ず `'<ブック名>'!<プロシージャ>` の形でブック名を付ける（ブック名中の `'` は `''` に重ねる）。付けないと、別のブックがアクティブなときにエラー 1004 になる。
* `Application.Run` はモーダルダイアログが出ている間は戻らない。呼び出し側は §14.10.6 のダイアログ監視を必ず併用し、ダイアログの内容を報告する。

### 14.6 Bridge による適用処理

* **標準モジュール・既存クラスモジュール:** `CodeModule.DeleteLines 1, CountOfLines` → `CodeModule.AddFromString 本文`。`VBComponents.Remove` → `Import` は使わない（`Remove` はマクロ終了まで遅延されるため、同名で再作成すると `Module11` のような別名になる罠がある）。
* **新規モジュールも必ず全行を削除してから書く:** VBE のオプション「変数の宣言を強制する」が有効な PC では、`VBComponents.Add` で作ったモジュールに最初から `Option Explicit` が 1 行入っている（無効な PC では空）。そのまま `AddFromString` すると生成コード側の `Option Explicit` と重複してコンパイル エラーになるため、新規・既存を問わず `CountOfLines > 0` なら `DeleteLines` してから書き込む。ユーザーの VBE 設定には依存しない。
* **新規クラス・属性変更を伴うクラス:** `AddFromString` では `Attribute` 行（`VB_PredeclaredId` 等）を設定できないため、この場合に限り一時ファイル経由で `Import` する。一時ファイルは **CRLF**・**実行中の Windows の ANSI コードページ**（`GetACP()`。`target.codepage` ではない）で書く。LF のみの `.cls` は標準モジュールとして取り込まれ、コードページが違うと文字化けする（§13.7）。
* **`Import` / `Remove` の結果を確認する:** `Import` 後のコンポーネント名が要求した名前と一致しない場合（同名のコンポーネントが残っていて `Foo1` になった等）は、取り込んだコンポーネントを削除して失敗として扱う。`Remove` が失敗した場合は `Import` に進まない。
* **ドキュメントモジュール（`ThisWorkbook`, シート）:** コードモジュールを直接書き換える。RAS 管理領域は `'--- RAS BEGIN ---` / `'--- RAS END ---` で囲み、それ以外の手書き部分は保持する。
* **差分適用:** 内容ハッシュが Excel 側の現在のハッシュ（Bridge が `RtlComputeCrc32` で計算しキャッシュ）と同じモジュールはスキップする。
* **一括性:** 1回の同期で複数モジュールを更新する場合、適用前の本文を Bridge が保持し、途中で失敗した場合は元に戻す。
* **実行中・中断モード中は適用しない:** マクロ実行中やブレークモード中にコードを書き換えると状態リセットや異常終了を招くため、Bridge はアイドル状態になるまで適用を保留し、ステータスを「保留」として返す。
* **適用後のコンパイル確認（任意）:** `ras.toml` の `sync.compileCheck = true` の場合、Bridge が VBE の「VBAProject のコンパイル」を実行し、エラーがあれば該当箇所をソースマップで RAS 側の位置に変換して VS Code に診断として表示する。
  * コンパイルのコマンドは VBE の**メニュー バー**から `FindControl(Id:=578)` で探す（「デバッグ」ツール バーにあるとは限らない）。コントロールが無効（`Enabled = False`）なのは「コンパイル済み」の意味であり、失敗ではない。
  * コンパイル エラーは VBE のモーダルダイアログとして出るため、実行は §14.10.6 のダイアログ監視の下で行う。ダイアログを閉じる前に VBE の選択位置（`ActiveCodePane.GetSelection`）からモジュールと行を読み取る。読み取れなくても、ダイアログを閉じる処理と結果の分類は変えない。
* **状態リセットの通知:** コード書き換えで VBA の状態（グローバル変数）がリセットされる旨をステータスに表示する。`@init`（§12.3）が次回のマクロ入口で再初期化する。

### 14.7 Excel → ローカル方向（`--pull` / 自動）

* **RAS 管理モジュール:** 一方向（RAS → Excel）のみ。生成コードの先頭に `' RAS-GENERATED hash=<CRC32>` を記録し、VBE 上で編集された場合（ハッシュ不一致）は**ドリフト**として警告し、「RAS で上書き」「差分表示」「VBE 側の編集を `.vba.patch` として保存」を選択させる。逆トランスパイルは行わない。
* **非管理モジュール（手書き VBA・ユーザーフォーム）:** 双方向同期する。`vba/` ディレクトリに `.bas` / `.cls` / `.frm` + `.frx` として保存する。Bridge は1〜2秒間隔で各モジュールの CRC32 を計算し、変化したものだけを共有メモリ経由で送り返す（ユーザーフォームのデザイナ情報 `.frx` はバイナリのため、変化検出時のみ一時ファイル経由で `Export` する）。ブック保存時（`Application.WorkbookBeforeSave`）にも同期する。
* **RAS からの非管理モジュール参照:** `vba/` 内の VBA モジュールの公開シグネチャを解析し、RAS から型付きで呼び出せるようにする（`import { Foo } from "vba:Module1";`）。

> **実装メモ（現行版）:** 非管理モジュールの双方向同期（`vba/` への保存と変化の検出）と `vba:` インポートは未実装。代わりに、同期のたびにブックの手書きの VBA を RAS に変換してプロジェクトに取り込み、RAS 管理に移す（§14.7.1）。

#### 14.7.1 手書き VBA の取り込み（RAS 管理への移行）（v0.9で追加）

ライブ同期（`ras sync`、`ras watch`、VS Code の言語サーバー）と `ras pack` は、書き込む前にブックの VBA プロジェクトを読み、RAS がまだ管理していない手書きのコードを RAS に変換してプロジェクトに取り込む。既存のブックをそのまま RAS のプロジェクトに移すための機能で、`ras.toml` の `[sync] pull-unmanaged = false` で無効にできる。

| ブックにあるもの | 取り込み先 | 条件 |
|---|---|---|
| 標準モジュール `Module1` | `src/Module1.ras` | `RAS-GENERATED` 行が無く、コードがある。プロジェクトに同名のモジュール（ファイル）が無い |
| シート `Sheet1`・`ThisWorkbook` の手書き部分（RAS 管理領域の外） | `src/Sheet1Events.ras`・`src/ThisWorkbookEvents.ras` | コードがある（`Option` 文・コメントだけのものは対象外） |
| VBE で作ったユーザーフォーム `frmInput` | `forms/frmInput.form.json`・`forms/frmInput.ras`（§14.11.6 と同じ） | 目印の行（`' RAS-FORM`）が無い。`forms/` に同名のファイルが無い |

* クラスモジュールは取り込まない（そのまま残る）。
* **変換:** `ras form pull` と同じ VBA → RAS の変換（§14.11.6）を使う。標準モジュールでは、`Public`（または省略）のプロシージャ・変数・定数・列挙型を `export` し、`Private` のものはそのモジュールの中だけのものにする。取り込むモジュール同士の参照は `import` にする（`Total(1)` → `import { Total } from "./Module1"`、`Module1.Total(1)` → `import * as Module1 from "./Module1"`）。
* **シート・`ThisWorkbook`:** イベント プロシージャ（`Worksheet_Change`、`Workbook_Open` など）は `@on("Sheet1.Change")` / `@on("Workbook.Open")` を付けた関数（§12.2）にする。`Me` と、修飾なしで使われたシート（ブック）のメンバー（`Range("A1")`、`Cells(1, 1)`、`Name`）は、`Sheet1.range("A1")` / `thisWorkbook.name` のようにシートのコード名（`thisWorkbook`）を付ける（VBA ではシートモジュールの中の `Range` はそのシートを指すため）。`Public` のプロシージャは `export` し、他のモジュールの `Sheet1.Foo` は `Sheet1Events.Foo` にする。`@on` が扱わないイベント（`Workbook_SheetPivotTableUpdate` など）は変換できないもの（TODO）として扱う。
* **モジュール名:** 取り込んだモジュールの VBA 名はファイル名から決まる（§12.1）。`mod_Util` のように `_` を含む名前は `ModUtil` になり、ブックの `mod_Util` は置き換えるときに削除される。ほかのモジュールやシートと名前が重なる場合は取り込まない（メッセージを表示する）。
* **元の VBA:** `vba/<名前>.bas`（シートは `.cls`、フォームは `.frm`）に、取り込んだ時点のコード（UTF-8・LF。シートは手書き部分だけ）を保存する。
* **フォームのデザイン:** ライブ同期では `VBComponent.Export` で書き出した `.frm` / `.frx` を読む（Excel が保存していない変更も含む）。`ras pack` ではブックのファイルから読む。

**置き換え:** 取り込んだファイルの 1 行目には目印 `// RAS-IMPORTED <名前>` を書く。目印のあるファイルについて、同期は次のように扱う。

| 取り込んだファイル | ブックの手書きのコード |
|---|---|
| `// TODO(VBA):` の行が無い | そのファイルから生成したコードで置き換える（上書きの指定は不要）。標準モジュール・フォームは置き換え、シートは手書き部分を削除して（`Attribute` 行・`Option` 文は残す）RAS 管理領域を書く |
| `// TODO(VBA):` の行がある | 置き換えずにそのまま残す。そのファイルから生成するモジュール（シートの場合は RAS 管理領域も）は書き込まない。同期の結果に「置き換えていないもの」として表示する |

* 変換できなかった行（§14.11.6）が残ったまま置き換えると、その処理がブックから失われるため、`// TODO(VBA):` の行がある間は置き換えない。直して `// TODO(VBA):` の行を消すと、次の同期で置き換える。
* 置き換えが成功したら、目印の行を通常のコメント（`// ブックの VBA（<名前>）から変換したファイルです。…`）に書き換える。以降は通常の RAS モジュールとして扱う。行数は変えない（行番号付きの生成コードが変わらないように）。
* 取り込んだファイルはプロジェクトのビルドに含まれる。変換結果にコンパイル エラーがあると、ほかのビルド エラーと同じく同期しない（ブックは変わらない）。
* 目印が無く、プロジェクトの同名のモジュールとブックの手書きのモジュールが重なる場合は、従来どおりエラーにする（上書きを指定すると置き換える。§14.9.3）。
* `ras pack` は、ブックが Excel で開かれているとき（オーナーファイルがあるとき）と `--output` を指定したときは取り込まない。

### 14.8 ソースマップとデバッグ

* 生成 VBA の各文と RAS ソース位置の対応を `build/*.rasmap` に保存する。
* `--erl` オプション（デバッグビルド既定）で生成 VBA に行番号を付与し、実行時エラーの `Erl` から RAS ソース位置を復元して表示する。
* VS Code 拡張の「生成 VBA を表示」では、カーソル位置に対応する VBA 行をハイライトする。
* ブレークポイントを使ったデバッグは、VBA ではなく JavaScript ランタイム上で行う（§14.10）。

### 14.9 オフライン書き込み（Excel 非起動）

Excel を起動せず、ブックファイル内の VBA プロジェクトを ras が直接読み書きする。初版から対応する。CI でのビルド・配布物の生成、Excel が無い環境での作業、ブックを閉じている間の反映に使う。

#### 14.9.1 対象ファイル

| 形式 | 対応 |
|---|---|
| `.xlsm` / `.xltm` / `.xlam` | 対応（OPC zip 内の `xl/vbaProject.bin`） |
| `.xlsb` | 対応（同上） |
| `.xlsx` | `ras pack` 時に `.xlsm` へ変換して出力（`--convert`）。元ファイルは書き換えない |
| `.xls`（BIFF8） | 対象外 |

#### 14.9.2 ファイル構造と書き込み手順

`vbaProject.bin` は複合ファイル（CFB）で、[MS-OVBA] に従い以下を書き換える。

1. **`VBA/<モジュール>` ストリーム:** モジュールソースを target のコードページ（§13.7）でエンコードし、MS-OVBA の圧縮形式で格納する。P-code（パフォーマンスキャッシュ）は書き込まず、ストリームはソースのみとする（`MODULEOFFSET = 0`）。
2. **`VBA/_VBA_PROJECT` ストリーム:** パフォーマンスキャッシュ無しを示す最小形（7 バイトヘッダ、Version = `0xFFFF`）に置き換える。これにより Office は次回オープン時にソースから再コンパイルし、ビルド番号の異なる Office 間（32bit / 64bit 含む）でも P-code 不整合が起きない。
3. **`VBA/dir` ストリーム（圧縮）:** `PROJECTCODEPAGE`、モジュールレコード（名前・Unicode 名・ストリーム名・種別・`MODULEOFFSET`・`MODULEPRIVATE` 等）を、追加・削除・名前変更に合わせて再生成する。
4. **`PROJECT` ストリーム:** `Module=` / `Class=` / `Document=` / `BaseClass=` 行と `[Workspace]` セクションを再生成する。`ID`・`CMG` / `DPB` / `GC`（保護情報）・`HelpFile` 等の既存値は保持する。
5. **`PROJECTwm` ストリーム:** モジュール名の MBCS / Unicode 対応表を再生成する。
6. **クラス属性:** `VB_PredeclaredId` 等の `Attribute` 行はソース先頭に含めて格納する（ライブ同期で `Import` が必要だった制約はオフラインでは無い）。
7. **ドキュメントモジュール:** `ThisWorkbook` とシートモジュールは、`xl/workbook.xml` の `workbookPr/@codeName` と各シートの `sheetPr/@codeName` に対応するもののみ書き換える（RAS 管理領域 `'--- RAS BEGIN/END ---` の扱いは §14.6 と同じ）。RAS のソースに存在するがブックに無いシートのコード名はエラーとする（シートの新規作成は行わない）。
8. **ユーザーフォーム・その他のストレージ:** 非管理のフォーム（デザイナストレージ）は `vba/` の `.frm` / `.frx` から再構成する。`forms/` のフォームは FormSpec からデザイナストレージを生成して書き込む（§14.11）。変更が無いストレージはバイト単位でそのまま保持する。
9. **OPC パッケージ:** `vbaProject.bin` のパート・リレーション・`[Content_Types].xml`（`application/vnd.ms-office.vbaProject`、ブック本体の content type を macroEnabled 版に）を整合させる。それ以外のパートは zip エントリを再圧縮せずそのままコピーする。

#### 14.9.3 安全性

* **原子的な書き込み:** 同じディレクトリの一時ファイルに書き出し、検証（後述）後に置き換える。置き換え前のファイルを `build/backup/` に保持する（世代数は `ras.toml` で指定）。
* **ロック検出:** 対象ブックが Excel で開かれている（ファイルがロックされている、または ROT に存在する）場合はオフライン書き込みを行わず、ライブ同期（§14.1〜14.8）に切り替える。`--force-offline` は用意しない。
* **書き込み後の検証:** 書き出した `vbaProject.bin` を ras 自身で再度読み込み、全モジュールのソースが一致することを確認する。
* **デジタル署名:** VBA プロジェクトが署名されている場合、書き換えで署名は無効になる。ras は警告を出したうえで署名パート（`vbaProjectSignature*.bin`）を削除する。`ras.toml` の `pack.sign` で署名用証明書を指定した場合は再署名する。
* **パスワード保護されたプロジェクト:** 保護（`DPB`）されている場合は書き込み・読み出しともエラーとする。
* **オーナーファイルの残骸:** ロック検出はオーナーファイル（`~$Book.xlsm`）の有無も見る。Excel が異常終了するとこのファイルが残り、開いていないのに「開かれている」と判定され続けるため、エラーメッセージで該当ファイルの削除を案内する。
* **手書きモジュールを黙って置き換えない:** ブックに、RAS が生成したものではない（`RAS-GENERATED` 行を持たない）同名の標準モジュール・クラス・フォームがある場合はエラーとし、何も書き込まない（名前の比較は大文字小文字を区別しない）。`--overwrite` で置き換えを許可する。
* **VBE でしか見つからないコンパイル エラーを事前に検出する:** オフライン書き込みはコンパイルを行わないため、ブックを開いて初めて VBE のダイアログで判明する種類の誤りは、書き込み前に検査する。ドキュメントモジュールで、RAS 管理領域が生成するプロシージャ（`Workbook_Open`、`Worksheet_Change` など）と同名のプロシージャが手書き部分にある場合（「名前が適切ではありません」になる）はエラーとし、何も書き込まない。
* **ドキュメントモジュールに `Option` 文を出力しない:** 手書き部分に `Option Explicit` があるかどうかは、そのモジュールを作ったときの VBE の設定（「変数の宣言を強制する」）次第で、ブックごと・シートごとに異なる。RAS 管理領域は宣言部ではなくモジュールの末尾に置かれるので、`Option` 文もモジュールレベルの宣言も含めず、プロシージャだけで構成する（生成コードは `Option Explicit` の有無にかかわらずコンパイルできる形にする）。

#### 14.9.4 オフライン読み出し（`ras unpack`）

* `VBA/<モジュール>` ストリームのソース部分を展開し、コードページから UTF-8 に変換して取り出す。非管理モジュール・フォームは `vba/` に、RAS 管理モジュールはドリフト検出（§14.7 と同じ `RAS-GENERATED hash`）にのみ使う。
* ソース部分が P-code と一致しない「VBA stomping」状態のファイルを検出した場合は警告する（ソース側を正とする）。
* `vba/` のファイルは UTF-8・LF で保存する（差分を取りやすくするため）。VBE のエクスポート形式（ANSI コードページ・CRLF）ではないので、VBE の「ファイルのインポート」には使えない。

---

### 14.10 デバッグ実行（JavaScript ランタイム）

VS Code 上でブレークポイント・ステップ実行・変数の確認を行うための、VBA とは別の実行経路。RAS のソースを JavaScript に変換して Node.js 上で動かし、Excel の操作が必要な箇所だけを COM 経由で実際の Excel に対して行う。

#### 14.10.1 方針

* **デバッガは自作しない:** 出力は CommonJS モジュールと Source Map（v3）であり、デバッガ本体は VS Code 組み込みの JavaScript デバッガ（js-debug）をそのまま使う。
* **書いたとおりに動かす:** インライン展開・ループ融合・一括転送・`DoEvents` 自動挿入（§8〜§10.4）は適用しない。RAS の 1 文が JavaScript の 1 文に対応し、ローカル変数名もそのまま残る。
* **VBA ビルドの代わりではない:** 意味論は VBA ビルドに合わせるが（§14.10.4）、実行タイミング・最適化後の挙動・Excel 起点の呼び出しは再現しない（§14.10.5）。実機で起きたエラーの位置は従来どおり `Erl` とソースマップ（§14.8）で特定する。

#### 14.10.2 出力（`ras build --target js`）

`build/debug/` に次のファイルを出力する（ビルドのたびに作り直す）。

| ファイル | 内容 |
|---|---|
| `<モジュール>.js` / `<モジュール>.js.map` | モジュールごとの JavaScript と Source Map。マッピングは文単位で、`sources` は `.ras` ファイルへの相対パス |
| `@forms/<フォーム名>.js` / `.js.map` | フォームのコードビハインド（`forms/<フォーム名>.ras`）。§14.10.7 |
| `ras-runtime.js` | ランタイム（`$rt`）。数値変換・オーバーフロー検査・文字列関数・エラーオブジェクト・VBA 標準関数・Excel へのアクセス |
| `ras-forms.js` | ユーザーフォームのランタイム（MSForms のオブジェクトとイベント、フォームを表示するホストとの通信）。§14.10.7 |
| `ras-main.js` | ランチャー。`node --enable-source-maps build/debug/ras-main.js <モジュール名.関数名>` で実行する（`--list` で実行可能な関数を JSON で出力） |
| `ras-debug.json` | 実行可能な関数の一覧とフォームの一覧（VS Code 拡張が読む） |

実行できるのは、引数なしで呼べるトップレベル関数と `Main`（トップレベルの文）、およびフォーム（`<フォーム名>.show` で、フォームを表示する。§14.10.7）。実行前にモジュール初期化と `@init` 関数を一度だけ実行する（VBA の `RAS_EnsureInit` と同じ）。

#### 14.10.3 Excel との接続

* ランタイムは Node アドオン `ras_com.node`（`crates/ras-com`、Windows 専用）を介して `IDispatch` を遅延バインディングで呼び出す。生成コードは `$rt.get(obj, "Value")` / `$rt.put(...)` / `$rt.call(obj, "Cells", [r, 3])` のように、プロパティ取得・設定・メソッド呼び出しを明示的に区別して呼ぶ。
* 接続先は §14.3 と同じ規則で決める: `ras.toml` の `workbook` を ROT から探し、開いている Excel インスタンスに接続する。開かれていなければ起動中の Excel で開く。Excel が起動していなければ、ブックをダブルクリックと同じ方法（`ShellExecute`）で開き、ROT に現れるまで最大 60 秒待つ。`CreateObject("Excel.Application")` で起動したオートメーション用の Excel は使わない（VBA プロジェクトへのアクセスなどが通常の起動と異なり、実行後に残ったインスタンスでは同期や「Excelで実行」ができなくなるため）。`workbook` が未指定ならアクティブブックを使う。
* ROT に載らない開き方がある（アドインとして読み込まれた `.xlam`、OneDrive / SharePoint の同期フォルダーにあり `FullName` が URL になるブック）。そのため `Workbooks.Open` の前に、接続した Excel の `Workbooks(<ファイル名>)` を直接引き、同じブック（フルパスが一致、または `FullName` が URL）ならそれを使う。同名の別のブックが開いている場合は、Excel のダイアログを待たずにエラーにする。
* `ThisWorkbook` はそのブック、シートのコード名（`Sheet1` など）は `CodeName` が一致するシート、その他のグローバル（`activeSheet`、`range(...)` など）は `Application` のメンバーとして解決する。メンバーはタイプ ライブラリの綴りで呼び出す（COM の名前解決は大文字小文字を区別しない）。
* Excel がセル編集中・モーダルダイアログ表示中で呼び出しを拒否した場合（`RPC_E_CALL_REJECTED` 等）は、最大 60 秒までバックオフしながら再試行する。待っている間の通知は §14.10.6。
* Excel のプロセスが無くなった場合（`RPC_S_SERVER_UNAVAILABLE` / `RPC_E_DISCONNECTED` 等）は、接続が切れた旨のエラーにする。
* 値の対応: `Empty` / `Null` / `Nothing` → `null`、日付 → 日付オブジェクト（シリアル値を保持）、セルのエラー値 → エラー値オブジェクト、`Range.Value` の 2 次元配列 → 0 始まりの行配列。
* VS Code 拡張は VS Code に同梱のランタイムを Node モード（`ELECTRON_RUN_AS_NODE`）で起動して使うため、Node.js を別途インストールする必要はない。組み込みデバッガの `launch` はこの環境変数を取り除いてしまうので、拡張がプロセスを起動し、デバッガは `attach` で接続する。

#### 14.10.4 VBA ビルドと合わせている意味論

* 整数演算のオーバーフロー（エラー 6）、ゼロ除算（エラー 11）、配列の範囲外アクセス（エラー 9）は実行時エラーになる。整数の除算・`as` による変換はゼロ方向への切り捨て。
* `any`（`Variant`）から型付き変数への代入は変換する（空セル → `0` / `""`、数値文字列 → 数値、変換できなければエラー 13、整数への変換は銀行丸め）。
* 配列と構造体（`struct`・無名のオブジェクト型・タプル）は値型（代入でコピー、関数には参照渡し）。`byref` 引数は呼び出し元の変数を書き換える。
* 実数の文字列化は有効 15 桁（`0.1 + 0.2` は `0.3`）。`Math.round` は四捨五入、`trim` は半角空白のみ除去。
* `try { } catch { resume; }` は文ごとにエラーを無視して次の文へ進む。`catch (e)` の `e` は `code` / `description` / `source` / `rasFile` / `rasLine` を持つ。
* `@fast` と `fastMode` / `suspend*` / `status` は実際の `Application` の状態を変更し、終了時（異常終了を含む）に元へ戻す。

#### 14.10.5 再現しないもの・制限

* **Excel 起点の呼び出し:** `@on` のイベントハンドラ（ブック・シート）、`@udf`、リボンコールバック、手書き VBA から RAS 関数を呼ぶ経路は起動できない（RAS から手書き VBA を呼ぶのは `Application.Run` 経由で可能）。ユーザーフォームは Excel ではなく VS Code に表示して動かす（§14.10.7）。
* **数値の `typeof`:** JavaScript の数値は `Long` と `Double` を区別しない。`typeof` は、静的な型に数値型が 1 つだけ含まれるとき（`double`、`string | double` など）はその型名を返し、それ以外では整数を `"Long"`、それ以外を `"Double"` とする。そのため `int | double` のように数値型を複数含むユニオン型を `typeof` で絞り込むと、VBA ビルドと結果が異なることがある。
* **数値型の精度:** `long`・`currency`・`decimal` は倍精度浮動小数点数で表現する（`long` は 2^53 を超えると不正確、`decimal` の `0.1m + 0.2m == 0.3m` は成り立たない）。`float` は倍精度のまま計算する。
* **VBA 標準関数:** よく使うもの（文字列・変換・日付・数学・`MsgBox`・`CreateObject`・`Format` の代表的な書式など）は JavaScript で実装している。未対応の関数は呼び出した時点でエラーになる。`Format` の未対応の書式は Excel の `TEXT` 関数に委譲する。ロケール依存の表記（日付の既定書式など）は VBA と一致しないことがある。
* **`declare`（DLL 関数）:** `kernel32` の `Sleep` / `GetTickCount` / `GetTickCount64` / `GetCurrentProcessId` / `QueryPerformanceCounter` / `QueryPerformanceFrequency` のみ代替実装がある。それ以外は警告 `W1410` を出し、呼び出した時点でエラーになる。
* **速度:** Excel へのアクセスはプロセス間 COM 呼び出しになるため、セルを 1 つずつ読み書きするループは VBA ビルドより大幅に遅い。
* **実行中の Excel:** マクロ実行中の状態ではないため、停止中も実行中もユーザーが Excel を操作できる。

#### 14.10.6 Excel のダイアログの監視

COM 呼び出しは、Excel がモーダルダイアログを出している間は戻らない。デバッグ実行中に起きるのは主に次の場合で、Excel は VS Code の後ろにあるため、何もしなければプログラムが固まったように見える。

* セルへの書き込みで動いた手書きの `Worksheet_Change` 等が実行時エラーになった（「実行時エラー」ダイアログ）
* 手書き VBA にコンパイル エラーがあり、イベントや `Application.Run` をきっかけに VBE がダイアログを出した
* 手書き VBA の `MsgBox` / `InputBox`
* Excel 自身の確認（シートの削除、上書き保存、リンクの更新、読み取り専用で開くか、など）

`ras_com.node` は、呼び出しが 0.3 秒以上戻らない間だけ、別スレッドで Excel プロセスのトップレベルウィンドウを調べる（呼び出しが出ていないとき＝ブレークポイントで停止中などは何もしない）。

* **検出:** 対象プロセス（`Application.Hwnd` の持ち主）のウィンドウのうち、クラスが `#32770`（標準ダイアログ）・`NUIDialog`・`bosa_sdm_*`（Office 独自のダイアログ）で、有効なもの。表示状態（`IsWindowVisible`）は条件にしない: Excel がバックグラウンドにあると、実行時エラーのダイアログは作られても描画されないことがある（オーナーとボタンを持つ `#32770` は非表示でも対象にする）。
* **内容の取得:** タイトル、`Static` の文言、`Button` のキャプション・コントロール ID・有効状態を Win32 で読む。子ウィンドウを持たない Office 独自のダイアログは UI Automation で文言とボタン名を読む（読めなければタイトルだけ報告する）。
* **種類の判定はキャプションの文言に頼らない:** ボタンはコントロール ID（`IDOK` / `IDCANCEL` …）とアクセスキー（`&E`、日本語版の `終了(E)`）で識別する。アクセスキー C・E・D を持つ 3〜4 ボタンの `#32770` は VBA の実行時エラー。
* **通知:** 見つけたダイアログは 1 回だけ、`[ras] Excel が…` で始まる 1 行として標準エラーに出す（タイトル・文言・ボタン）。VS Code 拡張はこの行を通知として表示する。Excel が呼び出しを拒否し続けているのにダイアログが見つからない場合（セルの編集中など）は、3 秒後にその旨を 1 回出す。60 秒のタイムアウトで失敗するときは、開いているダイアログをエラーメッセージに含める。

ダイアログの扱いは `launch.json` の `dialogs`（または設定 `ras.debug.dialogs`、環境変数 `RAS_DIALOGS`）で選ぶ。

| 値 | 動作 |
|---|---|
| `notify`（既定） | 通知し、ダイアログを前面に出す（前面化が OS に拒否された場合はタスク バーが点滅する）。応答はユーザーが行う |
| `dismiss` | 通知したうえで、安全に決まる応答があればそれで閉じる（下記）。人が見ていない実行（エージェント・CI）向け |
| `off` | 監視しない |

`dismiss` が選ぶ応答:

| ダイアログ | 応答 |
|---|---|
| VBA の実行時エラー | 「終了」（アクセスキー E）。「デバッグ」は VBE を中断モードに残し、以後の COM 呼び出しが通らなくなるため選ばない |
| `IDCANCEL` / `IDCLOSE` のボタンがある | そのボタン |
| 応答が 1 つしかない（ヘルプを除く。情報メッセージ、コンパイル エラー） | そのボタン |
| 子ウィンドウを持たない Office 独自のダイアログ | `WM_CLOSE`（キャンセルと同じ扱いになる） |
| それ以外（「はい / いいえ」のように選択が必要なもの） | 閉じない。通知だけ行い、ユーザーの応答を待つ |

ボタンは `BM_CLICK` で押し、閉じなければ `WM_COMMAND` の通知で再試行する（最大 3 回）。値を入力して続行するような応答（`InputBox` への入力など）は行わない。ダイアログを閉じた結果、元の呼び出しはエラー（キャンセルされた操作のエラー 1004 など）で戻ることがある。

#### 14.10.7 ユーザーフォーム（v0.9で追加）

フォームのコードビハインド（§14.11.2）もデバッグ実行の対象になる。デバッグ実行のプロセスは Excel のフォームのイベントを受け取れないため、フォームは Excel には表示せず、**VS Code のパネルに表示する**。コードビハインドは他のモジュールと同じく JavaScript として動き、ブレークポイント・ステップ実行・変数の確認が使える。Excel の操作（セルの読み書きなど）は従来どおり実際の Excel に対して行う。

* **起動:** フォームのコードビハインドの先頭の CodeLens「デバッグ実行」、デザイナーのタイトルバーのボタン、または実行する関数の一覧の `<フォーム名>.show` から、そのフォームを表示して始める。他のモジュールの関数から `frmInput.show()` で表示することもできる。
* **インスタンス:** フォームは既定のインスタンスだけを持つ（RAS ではフォームを `new` できない）。フォーム名を最初に参照した時点でインスタンスを作り、モジュール変数を初期化して `Initialize` を起こす。`unload` の後に参照すると、新しいインスタンスを作る（VBA と同じ）。
* **表示:** `show()` はモーダル（生成するフォームの `ShowModal` の既定値）、`show(vbModeless)` はモードレスで表示する。モーダル表示はフォームが閉じられる（`hide` / `unload`）まで戻らない。モードレス表示はすぐに戻り、起動した関数が終わった後も、表示中のフォームがすべて閉じられるまでプロセスは終わらない。
* **ホストとの接続:** VS Code 拡張はデバッグ実行のたびに `127.0.0.1` の TCP ポートで待ち受け、その番号を環境変数 `RAS_FORM_HOST` でプロセスに渡す。ランタイムは最初にフォームを表示するときに接続する。生成コードは同期的に動くため、ユーザーの操作を待つ間はワーカースレッドがソケットを受け持ち、メインスレッドは `Atomics.wait` で短い間隔ごとに待つ（その間もデバッガの一時停止は効く）。通信は 1 行 1 つの JSON で、ランタイムがフォームの状態（`show` / `update` / `hide` / `unload`）を送り、拡張はユーザーの操作（フォーカスの移動・入力・クリック・選択・タブの切り替え・閉じるボタン）を返す。`RAS_FORM_HOST` が無い（VS Code の外で実行した）場合、フォームの表示はエラー 429 になる。
* **状態はランタイムが持つ:** コントロールの値・フォーカス・表示状態はランタイムが持ち、パネルはそれを描画するだけである。入力した文字はランタイムに渡してから `Change` を起こし、ハンドラが書き換えた値をパネルに戻す。ハンドラの実行中（ブレークポイントで止まっている間を含む）に続けて入力した文字は失われない（パネルは最後の入力をランタイムが処理するまで、それより古い値で入力欄を上書きしない）。
* **イベント:** 次のイベントを VBA と同じ順序で起こす。ハンドラの呼び出し規則（宣言順・再入の無視・`byref` の引数）は §14.11.2 と同じ。

| きっかけ | イベント |
|---|---|
| インスタンスの作成 | `UserForm.Initialize`（モジュール変数の初期化の後） |
| 表示 | `UserForm.Activate`、タブ順で最初のコントロールの `Enter` |
| フォーカスの移動（クリック・Tab キー・`setFocus()`） | 移る前のコントロールの `BeforeUpdate`（値が変わっていれば）・`AfterUpdate`・`Exit`、移った先の `Enter`。`BeforeUpdate` / `Exit` の `cancel` で取り消すと、フォーカスは移らず、そのクリックや入力も行われない |
| 入力（`TextBox` / `ComboBox`） | `Change`（1 文字ごと） |
| クリック | `CommandButton` などの `Click`。`CheckBox` / `OptionButton` / `ToggleButton` は値を変えて `Change` と `Click`（同じグループの `OptionButton` は `false` になり `Change`） |
| リストの選択（`ComboBox` / `ListBox`） | `Change`、`Click` |
| `SpinButton` の矢印 | 値を `smallChange` だけ動かして `Change`、`SpinUp` / `SpinDown` |
| `ScrollBar` / `MultiPage` / `TabStrip` の操作 | `Change` |
| ダブルクリック | `DblClick` |
| Enter / Esc キー | `default` / `cancel` が `true` の `CommandButton` の `Click` |
| 閉じるボタン（パネルの ✕、パネルのタブを閉じる） | `UserForm.QueryClose`（`closeMode` = `vbFormControlMenu`）。`cancel` に 0 以外を入れると閉じない（タブを閉じた場合はパネルを開き直す）。閉じるときは `UserForm.Terminate` |
| `unload(this)` | `UserForm.QueryClose`（`closeMode` = `vbFormCode`）、`UserForm.Terminate` |

  コードでプロパティを変えたときも、VBA と同じく `Change` などが起きる（`text`・`value`・`listIndex` の変更）。キー入力のイベント（`KeyDown` / `KeyPress` / `KeyUp`）・マウスのイベント（`MouseDown` など）・ドラッグ＆ドロップ・`Deactivate` は起こさない。
* **エラー:** ハンドラで捕捉されなかったエラーは、VBA ビルドの RAS エラーダイアログ（`RAS_EventError`）と同じく実行を止めない。エラーの内容と RAS の位置をデバッグ コンソールとフォームのパネルに表示し、次の操作を待つ。エラーで止めたい場合は、VS Code のブレークポイントの「捕捉された例外」を有効にする。
* **MSForms のオブジェクト:** フォーム（`caption`・`width` / `height`・`backColor`・`font`・`controls`・`activeControl`・`show` / `hide`）と組み込みのコントロール（§14.11.1）の主なプロパティとメソッドを JavaScript で実装している（`text` / `value`、`caption`、`enabled` / `visible` / `locked`、位置と大きさ、色とフォント、`ComboBox` / `ListBox` の `addItem` / `removeItem` / `clear` / `list` / `column` / `listIndex` / `listCount` / `selected` / `multiSelect`、`MultiPage` の `pages` / `selectedItem`、`TabStrip` の `tabs`、`Image` の `picture` と `loadPicture`）。初期値は `.form.json` から読む（`list` / `selectedIndex` を含む）。未実装のメンバーはエラー 438、`controls.add` などの未対応の操作はエラー 445 になる。
* **見た目:** パネルは Windows の既定の色（システム カラー）とフォントでフォームを描く。座標と大きさはポイント単位で、Excel のフォームとおおむね同じ配置になるが、コントロールの描画（枠の形、リストの行の高さ、`ComboBox` のドロップダウンなど）はブラウザーのものであり、Excel とは一致しない。
* **再現しないもの:** Excel と VS Code の間のモーダル（モーダル フォームの表示中も Excel や VS Code の他の部分は操作できる。デバッグ実行の他のフォームは操作できない）、`StartUpPosition`、フォームのサイズ変更とスクロール、`ControlSource` によるセルとの連動、カスタム コントロール。

### 14.11 ユーザーフォーム（`forms/`）

ユーザーフォームのデザイン（コントロールの配置とプロパティ）をテキストで管理し、Excel を起動せずにブックへ書き込む。ブックにある既存のフォームを取り出すこともできる。デザインの形式、生成処理、デザイナストレージの読み取りは [xlflow](https://github.com/harumiWeb/xlflow)（MIT License, Copyright (c) 2026 harumiWeb）の FormSpec と、その純 Go 実装を移植したもの（`crates/ras-form`）。

#### 14.11.1 ファイル

| ファイル | 内容 |
|---|---|
| `forms/<Name>.form.json` | デザイン。xlflow の FormSpec（schemaVersion 1）の JSON 形式。ファイル名は `form.name` と一致させる |
| `forms/<Name>.ras` | コードビハインド。RAS で書き、フォームのコードモジュールにコンパイルする（§14.11.2）。無ければコードなしとして扱う |
| `forms/<Name>.bas` | 以前の版のコードビハインド（VBA。コードペインの内容をそのまま UTF-8 で書いたもの）。`.ras` が無いときだけ読む。両方あれば `.ras` を使い、警告を出す。`ras form pull --force` で取り込み直すと `.ras` に置き換わる |
| `forms/assets/` など | `Image` コントロールの画像（BMP / JPEG / PNG / GIF / WebP / TIFF / ICO）。置き場所は任意で、FormSpec からプロジェクトのルート基準の相対パスで参照する |

```json
{
  "schemaVersion": 1,
  "kind": "ras.userform",
  "basis": "designer",
  "form": { "name": "frmInput", "caption": "入力", "width": 240, "height": 180 },
  "controls": [
    { "id": "logo", "name": "imgLogo", "type": "Image", "left": 12, "top": 12, "width": 32, "height": 32, "picture": { "path": "forms/assets/logo.bmp" } },
    { "id": "kind", "name": "cboKind", "type": "ComboBox", "left": 56, "top": 12, "list": ["通常", "お急ぎ"], "selectedIndex": 0 },
    { "id": "name", "name": "txtName", "type": "TextBox", "left": 12, "top": 56, "width": 150, "height": 18, "controlTipText": "氏名" },
    { "id": "ok", "name": "btnOK", "type": "CommandButton", "caption": "OK", "left": 12, "top": 84 }
  ]
}
```

* `kind` は `ras.userform`（`ras form new` が書く値）または `xlflow.userform`。xlflow の JSON 形式の FormSpec はそのまま使える（YAML 形式は対象外）。
* 座標・サイズはポイント単位で、親コンテナー基準。`form.width` / `form.height` はクライアント領域（MS-OFORMS の `DisplayedSize`）。
* `controls` は平らな配列で、入れ子は `parentId` で表す。コンテナーは `Frame`・`MultiPage`（子は `Page` のみ）・`Page`（親は `MultiPage` のみ）。`TabStrip` のタブは `tabs` 配列に書く。
* 組み込みのコントロール: `Label`・`TextBox`・`ComboBox`・`ListBox`・`CommandButton`・`CheckBox`・`OptionButton`・`ToggleButton`・`SpinButton`・`ScrollBar`・`Image`・`Frame`・`MultiPage`・`Page`・`TabStrip`。
* 共通のプロパティ（`left` / `top` / `width` / `height` / `tabIndex` / `zIndex` / `enabled` / `visible`）、種類ごとのプロパティ（`caption` / `text` / `value` / `selectedIndex` / `tabs` など）、見た目（`font` / `foreColor` / `backColor` / `textAlign` / `borderStyle` / `pictureAlignment` / `pictureSizeMode`）の詳細と適用できるコントロールは xlflow の [UserForm Specification](https://harumiweb.github.io/xlflow/reference/userform-spec) に従う。色は `{ "kind": "rgb", "value": "#203040" }`、`{ "kind": "system", "value": "buttonFace" }`、`{ "kind": "ole", "value": "0x8000000F" }` の形で書く。
* 検査結果のコード（`UFY001`、`UFV001`〜`UFV021`）は xlflow と同じ意味を持つ。メッセージは日本語。`UFV022`（警告）は RAS だけのもので、1 行の `TextBox` / `ComboBox` の `height` がフォントの 1 行分に足りないことを表す。MSForms は欄をフォントに合わせて広げないので、Excel では文字が欠け、端に線が残る。必要な高さはフォントの大きさと種類（MS UI Gothic は 1 倍、Meiryo UI は 1.33 倍、Meiryo は 1.6 倍など）から求めた 1 行の高さに、枠と余白の 4.5pt を足し、1 ピクセル（0.75pt）単位に切り上げたもの。フォントの指定が無いときは MS UI Gothic 9pt として求める。デザイナーは、編集で高さかフォントが変わった欄をこの高さまで自動で広げる。

xlflow の FormSpec との違い（RAS で使える範囲が広い点）:

| 項目 | RAS での扱い |
|---|---|
| `tag` / `controlTipText` | すべてのコントロールで使える（xlflow では `Page` 以外は `properties` に書く） |
| `ComboBox` / `ListBox` の `list` / `selectedIndex` | 書き込みの対象（§14.11.5）。xlflow では取得した値の記録にだけ使う |
| コントロールのプロパティ | 値が 1 つのプロパティを、VBA の同名のプロパティと同じ値で書ける（名前は先頭を小文字にしたもの）。xlflow では `unsupported`。<br>・`TextBox`: `locked`・`multiLine`・`wordWrap`・`autoSize`・`passwordChar`・`scrollBars`・`maxLength`・`enterKeyBehavior`・`tabKeyBehavior`・`enterFieldBehavior`・`autoTab`・`autoWordSelect`・`hideSelection`・`selectionMargin`・`integralHeight`・`dragBehavior`・`imeMode`・`backStyle`・`specialEffect`・`borderColor`<br>・`ComboBox`: `style`（0: DropDownCombo、2: DropDownList）・`showDropButtonWhen`・`dropButtonStyle`・`listRows`・`listWidth`・`listStyle`・`matchEntry`・`matchRequired`・`locked`・`maxLength`・`autoSize` などの編集のプロパティ・`imeMode`・`backStyle`・`specialEffect`・`borderColor`<br>・`ListBox`: `multiSelect`・`listStyle`・`matchEntry`・`locked`・`integralHeight`・`imeMode`・`specialEffect`・`borderColor`<br>・`CommandButton`: `default`・`cancel`・`takeFocusOnClick`・`locked`・`wordWrap`・`autoSize`・`backStyle`<br>・`CheckBox` / `OptionButton` / `ToggleButton`: `locked`・`wordWrap`・`autoSize`・`backStyle`、`CheckBox` / `OptionButton` は `alignment`・`specialEffect`（0 か 2）、`OptionButton` は `groupName`<br>・`Label`: `wordWrap`・`autoSize`・`backStyle`・`specialEffect`・`borderColor`<br>・`SpinButton` / `ScrollBar`: `min`・`max`・`smallChange`・`orientation`・`delay`、`ScrollBar` は `largeChange`・`proportionalThumb`。`value` は `min` と `max` の間<br>・`Image`: `specialEffect`・`borderColor`<br>・`Frame`: `specialEffect`・`borderColor`・`cycle`<br>・`TabStrip` / `MultiPage`: `style`（0: Tabs、1: Buttons、2: None）・`tabOrientation`・`multiRow`・`tabFixedWidth`・`tabFixedHeight`（`MultiPage` のものは、その中の隠れた `TabStrip` に書き込む）<br>・`Image`・`Frame`・`TabStrip` と `Page` 以外のすべて: `mousePointer`。フォーカスを受け取るコントロール: `tabStop`<br>長さ（`listWidth`・`tabFixedWidth`・`tabFixedHeight`）はポイント、`borderColor` は色、`passwordChar` は 1 文字の文字列 |
| 書かなかったプロパティ | Excel でコントロールを作ったときと同じ値で書き込む（`ComboBox` はドロップ ボタンを常に表示・`matchEntry` 1、`ListBox` は `matchEntry` 0・スクロール バーあり、`Frame` は `specialEffect` 3 など）。xlflow は `ComboBox` と `ListBox` を MS-OFORMS の既定値（ドロップ ボタンなし・`matchEntry` 2）のまま書く。取り込み（`ras form pull`）では、これと違う値のものだけを書き出す |
| 書けないプロパティ | `TripleState`（ファイル形式の仕様に格納場所が無い）、`ColumnCount` / `ColumnWidths` / `BoundColumn` / `TextColumn` / `ColumnHeads`（`list` が 1 列のため）、`ControlSource` / `RowSource`、画像（`Image` の `picture` 以外）、`MouseIcon`、フォームそのもののプロパティ（`startUpPosition` など）。取り込んだフォームでは、これらは今までどおり `unsupported` に残る |

**`Image` の画像（`picture`）**

* `"picture": { "path": "forms/assets/logo.bmp" }` で画像ファイルを指定する。パスはプロジェクトのルート（`ras.toml` のあるフォルダー）からの相対パスで、ルートの外（`..`、絶対パス、ドライブ指定）は指定できない。
* 指定できる形式は BMP・JPEG・PNG・GIF・WebP・TIFF・ICO。ファイルの大きさは 16 MiB まで。形式は拡張子ではなくファイルの内容で判別する。メタファイル（WMF / EMF）は使えない。
* デザイナストレージに埋め込むのは、MS-Forms が読める形であることを xlflow が実機の Excel で確認している 2 種類だけとする: BMP（40 バイトの `BITMAPINFOHEADER`、24 ビットまたは 32 ビット、無圧縮）と JPEG。この形のファイルは内容のまま `StdPicture` として埋め込む（1,600 万ピクセルまで）。
* それ以外の画像（PNG・GIF・WebP・TIFF・ICO、および 8 ビットや RLE 圧縮・V4 / V5 ヘッダーなどの BMP）は、書き込みのたびにデコードして 24 ビットの BMP に変換してから埋め込む。元のファイルは書き換えず、`picture.path` は元のファイルを指したままでよい。
  * 透過部分は、その `Image` の `backColor` に重ねて塗りつぶす（BMP には透過が無いため）。`backColor` がシステム カラーのときは Windows の既定の色（`buttonFace` は `#F0F0F0`）、指定が無いときは `buttonFace` として扱う。
  * アニメーション GIF は最初のコマ、複数の画像を持つ ICO は最も大きいものを使う。
  * 変換後の BMP は 16 MiB まで（おおよそ 550 万ピクセル）。これを超える画像はエラーとする（小さくするか、JPEG にする）。
  * 同じファイルと同じ背景色からは、常に同じ BMP ができる（デザインのハッシュが変わらない）。
* ブックを開くのに元のファイルは要らない。
* `"picture": { "remove": true }` は画像なしを表す（`picture` を書かない場合と同じ）。

#### 14.11.2 コードビハインド（`forms/<Name>.ras`）

フォームのイベント処理などのコードは RAS で書く。コンパイラは `forms/<Name>.ras` を `src/` のモジュールと一緒に解析し、標準モジュールではなく**フォームのコードモジュール**として出力する（`ras build` は `build/` に書き出さず、`ras pack` / `ras sync` がデザイナストレージと一緒にブックへ書き込む）。

```typescript
// forms/frmInput.ras
let count = 0;

@on("UserForm.Initialize")
function init() {
    this.caption = `入力（${count}）`;
}

@on("btnOK.Click")
function ok() {
    if (this.txtName.text == "") {
        msgBox("名前を入力してください");
        this.txtName.setFocus();
        return;
    }
    count++;
    this.hide();
}

@on("txtName.Exit")
function leave(cancel: MSForms.ReturnBoolean) {
    cancel.value = this.txtName.text.length > 20;
}

@on("UserForm.QueryClose")
function closing(byref cancel: short, closeMode: short) {
    if (closeMode == vbFormControlMenu) cancel = 1;
}

/** 入力された名前（他のモジュールから `frmInput.enteredName()` で呼べる）。 */
export function enteredName(): string {
    return this.txtName.text;
}
```

* **`this`:** コードビハインドの関数の中では、`this` がフォーム（VBA の `Me`）を表す。型はフォーム名の型で、`MSForms.UserForm` のメンバー（`caption`・`hide()`・`controls` など）と、`.form.json` のコントロールを持つ。コントロールは `this.<コントロール名>` で参照し、型は `MSForms.TextBox` などの組み込みコントロールの型（カスタム コントロールは `object`）。フォームにないメンバーはエラー（`E0202`）。
* **型定義:** Microsoft Forms 2.0 の型（`MSForms.UserForm`・`MSForms.Control`・各コントロール・`MSForms.ReturnBoolean` などのイベント引数・`fm…` 列挙型）と `load` / `unload` / `loadPicture`、`vbModeless` / `vbFormControlMenu` などの定数を同梱する（`stdlib/msforms.d.ras`）。名前の綴りは他の型定義と同じく camelCase（§15.3.1）。生成する VBA では `MSForms.TextBox` のように名前空間付きで宣言する。
* **イベント:** `@on("<コントロール名>.<イベント>")` を付けた関数がイベントを受け取る。フォーム自身のイベントは `UserForm.<イベント>`（`Initialize`・`Activate`・`QueryClose`・`Terminate` など）。コンパイラは VBE と同じ宣言のイベント プロシージャ（`Private Sub btnOK_Click()`、`Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)` など）をフォームのコードモジュールに生成し、そこからハンドラを呼ぶ。イベントと引数は xlflow が VBE で採取した一覧に従う（VS Code 拡張と共通）。
  * ドキュメントモジュールのイベント（§12.2）と同じく、同じイベントに複数のハンドラを書け（宣言順に呼ぶ）、ハンドラ実行中に自身が発火させたイベントは無視する（`@on("…", { reentrant: true })` で解除）。ハンドラで捕捉されなかったエラーは RAS のエラーダイアログで表示する。
  * ハンドラの引数はイベントの引数の先頭から順に受け取り、省略できる。`ByVal` のオブジェクト引数（`MSForms.ReturnBoolean` など）は `value` に代入して結果を返す。`QueryClose` の `Cancel` のような `ByRef` の数値引数を書き換えるときは、`byref cancel: short` と宣言する（イベント プロシージャの宣言と型を合わせるため `Integer` は `short`）。
  * エラー: コントロールがない（`E0613`）、そのコントロールにないイベント（`E0611`。ページのイベントは親の `MultiPage` で受け取る）、引数が多すぎる（`E0612`）、関数名が生成するイベント プロシージャ名（`btnOK_Click`）と同じ（`E0614`）。
* **変数:** トップレベルの `let` / `const` はフォームのモジュール変数（`Private`）になる。初期値は `UserForm_Initialize` の先頭で設定する（`Initialize` のハンドラが無くても生成する）。フォームのインスタンスごとに持つ。
* **公開:** `export function` はフォームの `Public` メンバーになり、他のモジュールからフォームの既定のインスタンス経由で `frmInput.enteredName()` のように呼べる。フォーム名はどこからでも参照できるグローバル（既定のインスタンス）で、`frmInput.show()` のように使う。`export` しない関数は `Private`。コードビハインドの関数はマクロ入口にならない（§12.4 のラップもしない）。
* **書けないもの（`E0630`・`E0632`）:** トップレベルの文、クラス・インターフェイス・構造体・`declare module`、`export` した変数・列挙型・型エイリアス、`@macro` / `@init` / `@udf` / `@ribbon`、拡張メソッド。共有したいものは `src/` のモジュールに置き、`import` して使う（コードビハインドは `import` できる。コードビハインドを `import` することはできない）。
* `.form.json` が無い `forms/<Name>.ras` はエラー（`E0631`）。
* デバッグ実行（§14.10）では、フォームを VS Code のパネルに表示してコードビハインドを動かす（§14.10.7）。Excel 本体でフォームを表示して動かすのは「Excel で実行」（§14.11.8）。

#### 14.11.3 コマンド

| コマンド | 動作 |
|---|---|
| `ras form new <Name>` | `forms/<Name>.form.json`（240×180、caption = 名前、コントロールなし）と、書き方の例をコメントで示した `forms/<Name>.ras` を作る。既にあれば何もしない |
| `ras form check` | プロジェクトをコンパイルし、すべてのフォームを検査して、ブックへの書き込みと同じ生成まで行う。エラーがあれば終了コード 1 |
| `ras form pull [<Name>…] [--book <file>] [--force]` | ブックにあるフォームを `forms/` に取り出す。コードは RAS に変換する（§14.11.6） |
| `ras check` / `ras build` | RAS ソース（コードビハインドを含む）と合わせてフォームも検査する（`ras build` はフォームを `build/` に出力しない） |
| `ras pack` | RAS 管理モジュールと合わせて、すべてのフォーム（デザイナストレージとコードモジュール）をブックに書き込む |
| `ras sync` | Excel で開いているブックへ、RAS 管理モジュールと合わせてフォームを反映する（§14.11.7） |
| `ras watch` | `forms/` の変更も監視する（ブックが開いていれば `ras sync`、閉じていれば `ras pack` と同じ書き込みを行う） |

#### 14.11.4 書き込み（`ras pack`）

1. デザイナストレージ（`<Name>/f`・`o`・`\x01CompObj`・`\x03VBFrame`、コンテナーごとの `i<ID>/` サブストレージ、`MultiPage` の `x`）を FormSpec から一から生成する。生成結果は xlflow の生成結果とバイト単位で一致する（`crates/ras-form/tests/golden.rs`）。
2. コードモジュール `VBA/<Name>` を作る。新しいフォームには `VB_Base = "0{GUID}{GUID}"`（新しい 2 つの GUID）を含むフォーム用の属性ヘッダーを付け、既存のフォームはヘッダーを保持してコードだけを置き換える。
3. コードの先頭行は目印の行 `' RAS-FORM forms/<Name>.form.json design=<ハッシュ> code=<ハッシュ>` とし、RAS が生成したフォームであることを表す。`design` はデザイナストレージ、`code` はその後ろに続くコードのハッシュ（16 進 16 桁）で、ライブ同期が変更の有無を判断するのに使う（§14.11.7）。目印の後ろにコードビハインドをコンパイルしたコードと、必要なら生成コード（§14.11.5）が続く。
4. `PROJECT` ストリームに `BaseClass=<Name>` 行を加え、Microsoft Forms 2.0 Object Library（`MSForms`）の参照が無ければ `dir` ストリームに追加する。
5. 同名のフォームが既にあれば、デザイナストレージを丸ごと置き換える。デザインもコードも変わっていなければ「変更なし」として扱う。
6. 次の場合はエラーとし、何も書き込まない: RAS モジュールと同じ名前、フォームではない同名のモジュールがある、目印の無い（VBE で作った）同名のフォームがある（`--overwrite` で置き換えを許可する）。

`ras unpack` は、目印のあるフォームを `vba/` に取り出さない。`forms/` から消したフォームは、ブックからは削除しない（VBE で削除する）。

#### 14.11.5 リストの項目（`list` / `selectedIndex`）

MS-Forms は `ComboBox` / `ListBox` の項目をデザイナストレージに保存しない（デザイン時に `AddItem` しても、ブックを保存して開き直すと消える）。そのため `list` と `selectedIndex` は、フォームの読み込み時に実行されるコードとしてコードモジュールに出力する。

```vb
' RAS-FORM forms/frmInput.form.json design=… code=…
Option Explicit
Option Compare Binary

Private Sub init()
    Me.txtName.Text = ""
End Sub

Private Sub UserForm_Initialize()
    RAS_FormInit '--- RAS ---
    On Error GoTo RAS_OnError
    init
    Exit Sub
RAS_OnError:
    RAS_EventError "frmInput.UserForm_Initialize", Err.Number, Err.Source, Err.Description
End Sub
'--- RAS BEGIN ---
Private Sub RAS_FormInit()
    With Me.cboKind
        .AddItem "通常"
        .AddItem "お急ぎ"
        .ListIndex = 0
    End With
End Sub
'--- RAS END ---
```

* 項目を設定するプロシージャ `RAS_FormInit` を、モジュール末尾の RAS 管理領域（§14.6 と同じ `'--- RAS BEGIN ---` / `'--- RAS END ---`）に生成する。宣言部の後ろに置くので、コードビハインドの `Option Explicit` やモジュールレベルの宣言と干渉しない。
* コードビハインドに `UserForm_Initialize` が無ければ、`RAS_FormInit` を呼ぶだけの `UserForm_Initialize` も管理領域に生成する。
* コンパイルしたコードに `UserForm_Initialize`（`@on("UserForm.Initialize")` のハンドラや変数の初期化があると生成される）があれば、その先頭に `RAS_FormInit '--- RAS ---` の 1 行を挿入する。以前の版の VBA のコードビハインド（`.bas`）では、コードビハインドのどこかで `RAS_FormInit` を呼んでいる場合は、挿入も `UserForm_Initialize` の生成もしない（項目を設定する時点をコードビハインドが決められる）。
* `selectedIndex` は `-1`（選択なし）か、`list` の範囲内の番号。`list` が無いコントロールでは `-1` 以外を指定できない（`UFV020`）。
* 項目の文字列のうち、制御文字とコードページで表せない文字は `ChrW(n)` で出力する。
* コードビハインドでは `RAS_FormInit` という名前のプロシージャと、管理領域の目印の行を書けない（エラー）。

#### 14.11.6 取り出し（`ras form pull`）

ブック（保存されているファイル）のフォームを読み、`forms/<Name>.form.json`・`forms/<Name>.ras` と画像ファイルを書き出す。Excel は使わない。VBE で作ったフォームを `forms/` の管理に移すとき、VBE でデザインを直したフォームを取り込み直すときに使う。

* 名前を省略するとブックのすべてのフォームが対象。`forms/` に同名のファイルが既にあるフォームは書き出さない（`--force` で置き換える）。
* **デザイン:** デザイナストレージを読み、xlflow の projection と同じ規則で FormSpec にする（`crates/ras-form/tests/snapshot.rs` で、Excel が保存したフォームについて xlflow の出力と一致することを確認している）。加えて、xlflow の projection が「表せないプロパティ」として扱うもののうち、生成側が書けるものを取り込む: `Tag` / `ControlTipText`（`tag` / `controlTipText`）、`GroupName`・`MaxLength`・`BorderColor`（`properties`）。キャプションが保存されていないコントロールは `caption: ""` とし、フォームのキャプションがフォームのレコードに無ければデザイナブロック（`VBFrame`）の `Caption` を使う。
* **画像:** `Image` の画像（BMP / JPEG として保存されているもの）を `forms/assets/<Name>.<コントロール名>.bmp|jpg` に書き出し、`picture.path` で参照する。Excel は JPEG を読み込んだ画像も BMP で保存するため、取り出した画像は多くの場合 BMP になる。書き込み時に変換された画像（PNG など）も、取り出すと変換後の BMP になる。
* **コード:** コードモジュールから `Attribute` 行を除いた VBA を RAS に変換して `forms/<Name>.ras` に書く（下記）。RAS が書き込んだフォームの場合は、目印の行・RAS 管理領域・挿入した `RAS_FormInit` の呼び出しも除いてから変換する。
  * RAS が書き込んだフォーム（目印の行がある）で、`forms/<Name>.ras` が既にあるときは、コードビハインドを書き換えない（そのファイルからコンパイルしたコードなので）。デザインだけを取り込み直す。VBE でコードを編集しても取り込まれない。
  * 以前の版の `forms/<Name>.bas` は、`.ras` を書いたときに削除する。
* **VBA から RAS への変換:** 構文を読み替えて RAS を書く（`ras_compiler::vba2ras`）。
  * イベント プロシージャ `<コントロール>_<イベント>` は `@on("<コントロール>.<イベント>")` を付けた関数（名前は `btnOKClick` のように付ける）に、`Me` は `this` に、コントロールの参照は `this.<コントロール名>` にする。既定メンバーで使われているコントロール（`txtName = "x"`、`If chkAll Then`）と `MSForms.ReturnBoolean` などの引数（`Cancel = True`）は `.value` を補う。
  * ホストの関数・メンバーは RAS の綴り（`MsgBox` → `msgBox`、`.ListIndex` → `.listIndex`）、型は RAS の型（`Long` → `int`、`Integer` → `int`（イベントの引数は `short`）、`Range` → `Excel.Range`、`TextBox` → `MSForms.TextBox`）にする。
  * `&` の連結はテンプレートリテラル、`/` は実数の除算（整数同士なら `as double` を補う）、`=` / `<>` は `==` / `!=`、`And` / `Or` / `Not` は `&&` / `||` / `!`（数値の演算は `&` / `|` / `~`）にする。
  * `If` / `Select Case`（`if` の連鎖にする）/ `For` / `For Each` / `Do … Loop` / `While … Wend` / `With`（対象を繰り返し書くか、一時的な `const` に入れる）/ `Exit` を変換する。`On Error Resume Next` … `On Error GoTo 0` は `try { … } catch { resume; }`、`On Error GoTo <ラベル>` … `Exit Sub` / `<ラベル>:` … は `try { … } catch (e) { … }`（ハンドラ内の `Err.Number` / `Err.Description` は `e.code` / `e.description`）にする。
  * `Function` の戻り値（`関数名 = 値`）は `return` かローカル変数 `result` にする。引数は `ByVal` が既定になるため、`ByVal` の無い引数はプロシージャ内で代入している場合だけ `byref` にする。
  * RAS が書き込んだフォームのコード（`.ras` の無いプロジェクトへ取り出したとき）は、生成したイベント プロシージャを `@on` に、変数の初期化（`RAS_ModInit`）を変数の初期値に戻す。行番号（`Erl` 用）は捨てる。
  * 変換できないもの（`GoTo` / 行ラベル、`ReDim`、大きさを指定した配列、`Like`、日付リテラル、`Static`、引数を取るプロパティへの代入、ユーザー定義型・`Declare`・`Public` 変数（フォーム内の変数にする）など）は、元の行を `// TODO(VBA): <理由>` のコメントとして残し、コマンドの出力に警告（ファイルと行）として表示する。変換結果は必ずしもそのままコンパイルできるとは限らない。`ras check` で確認して直す。
* **リストの項目:** ブックには保存されていないので取り出せない。`--force` で既存の `.form.json` を置き換えるときは、同名の `ComboBox` / `ListBox` の `list` / `selectedIndex` と、画像の `picture.path` を既存のファイルから引き継ぐ（そのパスのファイルを書き込み時と同じ規則で変換した結果が、ブックの画像と一致する場合。PNG などを指定していたフォームは、取り込み直しても PNG を指したままになる）。
* **取り込めないもの:** FormSpec で表せないプロパティ（`specialEffect`、`mousePointer`、スクロールバーの `min` / `max` など）は、コントロールごとに `unsupported` に名前を記録し、コマンドの出力に警告として表示する。これらは `forms/` から書き戻すと既定値になる。組み込み以外のコントロール（ActiveX）は `type: "Control"` として記録され、生成できない。BMP / JPEG 以外の画像（アイコン、メタファイル）は取り出せず、`unsupported` に `picture` が残る（この場合は `picture` に置き換える画像か `{ "remove": true }` を書くまで生成できない）。
* 取り出しただけではブックは変わらない。VBE で作ったフォームを `forms/` の内容で置き換えるには、`ras pack --overwrite`（または `ras sync --overwrite`）を実行する。
* `forms/` に無いフォームは、同期（`ras sync`・`ras pack` など）のときにも同じ変換で取り込まれ、置き換えられる（§14.7.1）。

#### 14.11.7 ライブ同期（`ras sync`、`ras watch`、VS Code の言語サーバー）

ブックが Excel で開かれているときは、§14.6 のモジュールと同じ同期の中でフォームも反映する（Windows のみ）。

| 状態 | 動作 |
|---|---|
| ブックにフォームが無い | 取り込む（作成） |
| 目印の行が同じ | 何もしない |
| `design` が同じで `code` が違う | `CodeModule` のコードだけを書き換える |
| `design` が違う | フォームを取り込み直す（更新） |
| 目印の無い同名のフォーム（VBE で作ったもの） | エラー。上書きを指定した場合は取り込み直す。同期で取り込んだフォーム（§14.7.1）は取り込み直す |

* **取り込み:** フォームを `.frm`（デザイナブロック・属性・コード）と `.frx`（24 バイトのヘッダーと、デザイナストレージを収めた複合ファイル）として一時フォルダーに書き出し、`VBComponents.Import` で読み込ませる。`.frx` の内容は `ras pack` が書くデザイナストレージと同じもの。
* **取り込み直し:** 既存のフォームを一時的な名前（`RasOld1` など）に変えてから新しいフォームを取り込み、同期全体が成功した後で古いフォームを削除する。途中で失敗した場合は、新しいフォームを削除して古いフォームの名前を戻す（§14.6 の一括性）。
* **読み込まれているフォーム:** 取り込み直すフォームとコードを書き換えるフォームは、書き込む前に閉じる。VBE で開いているデザイナーウィンドウを閉じ、読み込まれているインスタンス（`Hide` で隠しただけのフォーム、モードレスで表示中のフォーム）を `Unload` する（VBE は読み込まれているフォームの名前を変更できないため）。`Unload` はプロジェクトに一時的な標準モジュール（`RAS_SyncUnload`）を加えて `Application.Run` で実行し、終わったら削除する。フォームの `QueryClose` / `Terminate` のハンドラは実行される。`QueryClose` で閉じるのを取り消されたフォームがあれば、何も書き込まずにエラーにする。
* **変更の判断:** VBE はコードを整形して保持するため、保持されているコードと書き込んだコードは比較できない。目印の行に記録したハッシュで判断する。
* VBE でフォームのデザインやコードを直接編集しても目印の行は変わらないため、その編集は次に `forms/` 側が変わるまで残る（ドリフト検出は行わない）。VBE でのデザインの編集を残したい場合は、ブックを保存してから `ras form pull --force` で `forms/` に取り込む（コードの編集は取り込まれない。§14.11.6）。
* マクロの実行中・中断モード中は書き込まない（§14.6）。フォームを取り込み直すと、そのフォームのインスタンスとプロジェクトの状態はリセットされる。

#### 14.11.8 Excel での実行（v0.9で追加）

VS Code の「Excel で実行」（フォームのコードビハインドの先頭の CodeLens、デザイナーのタイトルバーのボタン、コマンド「フォームを Excel で実行」。Windows のみ）は、ブックへ同期（§14.11.7）してから、Excel 本体でフォームを表示する。

* `Application.Run` はフォームのコードモジュールのプロシージャを呼べず、`VBA.UserForms` はプロジェクト内のコードからしか使えない。そのため、プロジェクトにフォームがあるときは、ランタイム（`RAS_Runtime`）に次のプロシージャを出力し、`Application.Run "'<ブック名>'!RAS_Runtime.RAS_ShowForm", "<フォーム名>"` で呼ぶ。

  ```vb
  Public Sub RAS_ShowForm(ByVal name As String)
      VBA.UserForms.Add(name).Show
  End Sub
  ```

* 表示の前に Excel を前面に出す（Windows が前面の切り替えを許可しない場合は、タスク バーのボタンが点滅する）。
* モーダルのフォームは閉じられるまで呼び出しが戻らない。その間、同期（保存のたびの同期を含む）は待たされる。

#### 14.11.9 制限

* カスタム ActiveX コントロールは生成できない（エラーになる）。
* VBE でフォームのデザインやコードを編集しても `forms/` には自動では反映されない。デザインは `ras form pull --force` で取り込める。取り込まない限り、次に `forms/` 側を変更して書き込んだ時点で `forms/` の内容に置き換わる（ドリフト検出は行わない）。
* コードビハインドで扱えるのは組み込みのコントロールのイベントだけで、`WithEvents` で受け取るイベントや、コントロールの配列（同じハンドラを複数のコントロールに付けるクラス）は書けない。
* `ComboBox` / `ListBox` の項目は文字列の 1 列だけを扱う。複数列のリストや `RowSource` は、コードビハインドで設定する。
* ライブ同期でのフォームの取り込み（`VBComponents.Import`）は、Excel の無い環境では自動テストできない。`.frx` の構成は Excel が書き出したファイルと照合している（`crates/ras-form/tests/project.rs`）。

## 15. プロジェクト構成

### 15.1 ディレクトリ

```
myproject/
  ras.toml
  src/            *.ras（RAS 管理モジュール）
  vba/            非管理の VBA モジュール・ユーザーフォーム（双方向同期）
  forms/          ユーザーフォームのデザイン（*.form.json）とコードビハインド（*.ras）（§14.11）
  types/          追加の型定義（*.d.ras）
  build/          生成物（.bas / .cls / .rasmap）
  build/debug/    デバッグ実行用の JavaScript とソースマップ（§14.10）
```

### 15.2 `ras.toml`

```toml
[project]
name = "SalesReport"
workbook = "C:/work/SalesReport.xlsm"

[target]
host = "excel"          # 現状 "excel" のみ
vba6 = true             # VBA6(32bit Office 2007以前)用の Declare も出力
codepage = 932

[build]
opt-level = 1           # 0: 展開なし / 1: ループ内自動展開 / 2: 積極的
inline-threshold = 40
erl = true              # 行番号付与

[pack]
backup-generations = 5
# sign = "cert:CurrentUser/My/<thumbprint>"   # 再署名に使う証明書

[doevents]
enabled = true
interval-ms = 100       # DoEvents を呼ぶ最短間隔
loop-threshold = 10000  # 回数が定数のループで挿入対象とする最小回数

[sync]
enabled = true
compileCheck = true
shared-memory-mb = 64
pull-unmanaged = true   # 同期・ras pack のときに、ブックの手書きの VBA を RAS に変換して取り込む（§14.7.1）
```

### 15.3 ホスト型定義

* Excel オブジェクトモデル（`Excel.Range`, `Excel.Worksheet` 等）と Office・VBA 標準関数の型定義を同梱する。
* その他の COM ライブラリ（ADODB 等）は `ras typegen` でタイプライブラリから生成する。生成した型を使ったコードは既定で遅延バインディング（`Object` + `CreateObject`）で出力し、参照設定への依存を避ける（`@earlyBinding` で事前バインディングに切り替え）。

#### 15.3.1 名前の綴り（v0.9で追加）

Excel / VBA / Office のタイプ ライブラリは、メンバー名を PascalCase（`Cells`、`ScreenUpdating`）で定義している。RAS のソースでは、JavaScript の命名規則に合わせて次のように参照する。補完・クイック情報・診断メッセージもこの綴りで表示する。

| 対象 | RAS での綴り | 例 |
|---|---|---|
| プロパティ・メソッド | camelCase | `ws.cells(r, 1).value`、`application.screenUpdating`、`range.interior.color` |
| グローバルのオブジェクト・関数 | camelCase | `application`、`thisWorkbook`、`activeSheet`、`worksheets`、`msgBox(...)`、`left(s, 1)`、`debug.print(...)`、`err.number` |
| 型（インターフェイス・列挙型）・名前空間 | PascalCase のまま | `Excel.Worksheet`、`XlCalculation`、`VBA.ErrObject` |
| 列挙型の定数 | タイプ ライブラリのまま | `xlUp`、`vbCrLf`、`msoTrue` |

* **変換規則:** 先頭の大文字の並びを 1 語として小文字にする。次の語が続く場合、その語の頭文字は大文字のまま残す。`Value` → `value`、`VBProject` → `vbProject`、`RGB` → `rgb`、`UBound` → `uBound`、`CStr` → `cStr`、`FormulaR1C1` → `formulaR1C1`。
* **キーワードと重なるグローバル:** 変換後の名前が RAS のキーワードになるグローバルは、VBA の綴りのまま使う（`Switch(...)`）。メンバーは `.` の後に書くためキーワードと重なってよい（`range.delete()`、`ws.select()`）。
* **生成する VBA:** タイプ ライブラリの綴りで出力する（`ws.Cells(r, 1).Value`）。型のない値（`any` / `object`）のメンバーも、同梱の型定義に同じ名前のメンバーがあればその綴りで出力する（`activeSheet.cells(1, 1)` → `ActiveSheet.Cells(1, 1)`）。
* **VBA の綴りで書いた場合:** VBA は大文字小文字を区別しないため、`ws.Cells(r, 1).Value` と書いても同じメンバーとして解決する。ただし警告（`W0304`）を出す。`ras check --fix` とエディタのクイックフィックスが camelCase に書き換える。
* **プロジェクトが宣言した名前はそのまま:** シートのコード名（`Sheet1`）と `types/*.d.ras` で宣言した名前は、書かれたとおりに参照し、そのまま出力する。変換するのは同梱の型定義の名前だけである。
* **同梱の型定義の書き方:** `.d.ras` 自体はタイプ ライブラリの綴りで書く（`ras typegen` の出力と同じ形）。コレクションの要素の型は、型引数で表す（`interface Sheets<T>`）。型引数を省略したホスト型は、既定の型引数があればそれ（`Excel.Sheets` は `Excel.Sheets<Excel.Worksheet | Excel.Chart>`。§2.6.5）、なければ要素が `object` になる。タイプ ライブラリが Object / Variant を返すメンバーでも、値の型が決まっているものはユニオン型で書く（§2.6.5）。

---

## 16. エディタ支援（IntelliSense）とドキュメントコメント

### 16.1 ドキュメントコメント（JSDoc 形式）

ドキュメントコメントは JSDoc 形式（`/** ... */`）とする。

```typescript
/**
 * 指定列の最終行を返す。
 *
 * 空の列では `1` を返す。
 * @param col 列番号（1 始まり。Excel の列番号に合わせる）
 * @returns 最終行の行番号
 * @example
 * const last = Sheet1.lastRow(2);
 * @see {@link firstEmptyRow}
 */
export function lastRow(this: Excel.Worksheet, col: int = 1): int { ... }
```

#### 16.1.1 対象

関数・拡張メソッド・クラス・クラスメンバー・インターフェイス・`const` / `let`（モジュールレベル）・`declare` の外部関数・型定義（`.d.ras`）の宣言の直前に書いた `/** */` を、その宣言のドキュメントとして扱う。

#### 16.1.2 対応タグ

| タグ | 用途 |
|---|---|
| （本文） | 概要。1 段落目を要約として補完一覧に、全文をクイック情報に表示する。Markdown を使用可 |
| `@param 名前 説明` | 引数の説明。パラメーター情報で該当引数の説明として表示する |
| `@returns 説明` | 戻り値の説明 |
| `@template T 説明` | ジェネリクスの型パラメータの説明（§2.4） |
| `@throws 説明` | 送出しうるエラー |
| `@example` | 使用例（コードブロックとして表示） |
| `@remarks` | 補足 |
| `@deprecated 説明` | 非推奨。使用箇所に取り消し線と警告を表示する |
| `@see` / `{@link 名前}` | 関連項目へのリンク。`{@link}` はクイック情報内で定義へジャンプできるリンクになる |
| `@since` | 追加されたバージョン |

* JSDoc の型注釈（`@param {string} x` の `{string}`）は無視する。型はコード上の型注釈を正とする。
* `@param` の名前が実際の引数名と一致しない場合は警告する。
* `export` された宣言にドキュメントコメントが無い場合の警告は既定で無効（`ras.toml` の `[lint] require-docs = true` で有効化）。

#### 16.1.3 VBA への出力

* 関数・クラスのドキュメントコメントの要約は、生成 VBA の宣言の直前に `'` コメントとして出力する（コードページで表せない文字は §13.7 の規則で変換）。
* `Public` なプロシージャには `Attribute <名前>.VB_Description = "<要約>"` を出力し、VBE のオブジェクト ブラウザーにも説明が出るようにする。ただし、ライブ同期（§14.6）の `AddFromString` は `Attribute` 行を受け付けないため、ライブ同期では出力しない。`Import` 経由とオフライン書き込み（§14.9）では出力する。
* 通常のコメント（`//`, `/* */`）は、ソースマップで対応が取れる位置に `'` コメントとして出力する（`ras.toml` の `[build] keep-comments`、既定 `true`）。

### 16.2 IntelliSense 機能

VS Code の IntelliSense に相当する機能を、LSP の標準機能として `ras lsp` が提供する。

| 機能 | LSP メソッド | 内容 |
|---|---|---|
| メンバー一覧（補完） | `textDocument/completion`, `completionItem/resolve` | `.` の後に、レシーバの型のメンバー（RAS クラス、Excel 等の COM 型、配列・文字列の組み込みメンバー、コレクションの配列メソッド（§4.4））と適用できる拡張メソッドを表示する。COM 型のメンバーとグローバルは camelCase で表示する（§15.3.1）。`.` 以外の位置では、スコープ内の変数・関数・型・キーワード・スニペットを表示する。引数リストの中では名前付き引数の引数名も表示する（§16.2.1） |
| クイック情報（ホバー） | `textDocument/hover` | RAS 構文のシグネチャ、JSDoc（要約・引数・戻り値・例）、推論された型を表示する |
| パラメーター情報 | `textDocument/signatureHelp` | `(` と `,` で起動し、シグネチャと現在の引数をハイライトし、その引数の `@param` の説明を表示する |
| 定義へ移動 / 型定義へ移動 | `textDocument/definition`, `typeDefinition` | ジェネリクスや拡張メソッドも元の定義へ移動する |
| 参照の検索 | `textDocument/references` | |
| 名前の変更 | `textDocument/rename`, `prepareRename` | 変更後の名前が大文字小文字違いの衝突（§13.6）や VBA 予約語に当たる場合は拒否する |
| シンボル一覧 | `textDocument/documentSymbol`, `workspace/symbol` | アウトライン表示、シンボル検索 |
| インレイヒント | `textDocument/inlayHint` | 推論された型、推論されたジェネリクスの型引数、呼び出し側の `byref` 引数（`byref` と表示） |
| セマンティックトークン | `textDocument/semanticTokens` | TextMate 文法で区別できない `byref` 引数・拡張メソッド・COM メンバー等を色分けする |
| 診断 | `textDocument/publishDiagnostics` | §6 のリンタールール、同期時のコンパイルエラー（§14.6） |
| コードアクション | `textDocument/codeAction` | クイックフィックス（VBA の綴りで書かれたホストの名前を camelCase に直す（§15.3.1。その箇所だけ／ファイル内すべて）、未インポートの関数・拡張メソッドの `import` 追加、`byref` 引数へのリテラル渡しを一時変数に置き換える、縮小変換に `as` を追加する等） |
| CodeLens | `textDocument/codeLens` | 最適化の適用状況（インライン展開・一括転送・ループ融合・`DoEvents` 挿入）、「生成 VBA を表示」、単独で起動できる関数の「Excelで実行」「デバッグ実行」、フォームのコードビハインドの先頭の「Excelで実行」「デバッグ実行」（§14.10.7、§14.11.8） |
| フォーマット | `textDocument/formatting`, `rangeFormatting` | `ras fmt` と同じ |
| 折りたたみ | `textDocument/foldingRange` | |

#### 16.2.1 メンバー一覧の詳細

* **並び順:** 型のメンバーを先に、拡張メソッドを後に表示する。拡張メソッドにはアイコンと `(拡張)` の表示を付ける。
* **未インポートの候補:** プロジェクト内のまだインポートしていないモジュールの `export` された関数・拡張メソッドも候補に出し、確定時に `import` 文を自動で追加する。
* **`any` / `object` 型:** 静的型が無いため、メンバー一覧は表示しない（遅延バインディングになる旨をヒントで示す）。
* **列挙型の定数:** `xlUp` 等の COM 列挙型の定数は、期待される型が列挙型の引数位置で優先的に表示する。
* **名前付き引数:** 宣言のある関数・メソッド・コンストラクタ・拡張メソッド・Excel / VBA のメンバーの引数リストで、引数の先頭（`(` や `,` の直後）にいるとき、まだ指定していない引数の名前を `name:` の形で一覧の先頭に表示し、確定すると `name: ` を挿入する（§3.3）。位置で渡した引数・名前で渡した引数・可変長引数・拡張メソッドの `this` 引数は除く。名前付き引数の後ろでは引数名だけを表示する。`(` と `,` の入力でも一覧を開く（このときは引数名だけを表示する）。
* **ドキュメントの遅延読み込み:** 一覧の段階では名前と型だけを返し、JSDoc の本文は `completionItem/resolve` で返す。

#### 16.2.2 クイック情報の詳細

* 関数: RAS 構文のシグネチャ、JSDoc、`inline` か、自動展開の対象になったか。
* 変数: 推論された型。出力される VBA の型（例: `int` → `Long`）も併記する。
* 拡張メソッドの呼び出し: 定義モジュールと、出力される VBA の呼び出し形（例: `StrUtil.isBlank(name)`）。
* ジェネリクス: 推論された型引数と、出力される実装名（例: `maxOf__Long`）。
* COM 型のメンバー: 型ライブラリのヘルプ文字列と、Microsoft のリファレンスへのリンク（`@see`）。VBA での綴り（例: `// VBA: Worksheet.Cells`）も併記する。

#### 16.2.3 Excel 等の型定義

* 同梱する Excel / Office / VBA 標準の型定義（`.d.ras`、§15.3）には JSDoc を付ける。JSDoc の内容は、`ras typegen` でタイプライブラリのヘルプ文字列とメンバーの引数名を取り込み、Microsoft のリファレンスページへの `@see` リンクを付けて生成する。リファレンス本文は同梱しない（ライセンス上の理由）。
* `ras typegen` で生成した ADODB 等の型定義にも同じ規則を適用する。

#### 16.2.4 応答性能

* 入力中の不完全なコードでも補完できるよう、パーサーはエラー回復を行い、構文木（ロスレス CST）を常に作る（§17.1）。
* 解析はインクリメンタル計算（§17.1 の `salsa`）で、変更されたファイルと依存先だけを再計算する。
* 目標: 10 万行のプロジェクトで、補完・ホバー・パラメーター情報の応答 50ms 以内（目標値）。
* 最適化（§8〜§10）は IntelliSense の応答経路では実行しない。CodeLens 用の最適化結果はバックグラウンドで計算し、完了後に更新する。

#### 16.2.5 参照の検索・名前の変更の詳細

* **対象:** 関数・メソッド・アクセサ・拡張メソッド・`declare` 関数、クラス・インターフェイス・構造体・列挙型・型エイリアス、フィールド・構造体のフィールド・列挙値、モジュールレベルの変数・定数、ローカル変数・引数（ラムダの引数を含む）。宣言の位置からも、使用箇所からも起動できる。
* **参照に含めるもの:** 式中の名前とメンバー、型注釈・`new`・`as`・`implements` 等の型名、`import { 名前 }`、名前付き引数の引数名、構造体のオブジェクトリテラルのキー、`@onError(関数)` / `@fallback(関数)` の関数名。名前の変更では、引数のドキュメントコメントの `@param 名前` も書き換える。オブジェクトリテラルの省略形 `{ x }` は `{ y: x }`（フィールドの変更）/ `{ x: y }`（変数の変更）に書き換える。
* **まとめて変わるもの:** getter と setter、インターフェイスのメンバーとそれを実装するクラスの同名のメンバー（その先の実装も含む）。
* **インポート時の別名:** `import { a as b }` の `b` の使用箇所からは名前の変更を起動できない（元の宣言から変更すると `import { a2 as b }` になり、`b` の使用箇所はそのまま）。
* **名前を変更できないもの:** 型定義（`.d.ras`）と同梱のライブラリの宣言（Excel / VBA の名前）、フォームのコントロール（`forms/<名前>.form.json` で変更する）、`@alias` の無い `declare` 関数（DLL の関数名）、コンストラクタ、`this`、無名のオブジェクト型・タプルのフィールド。
* **拒否する新しい名前:** 識別子でないもの、RAS のキーワード・基本型名、VBA の予約語（§13.6）、`RAS_` で始まる名前、プロジェクト内の他の識別子と大文字小文字だけが異なる名前（§13.6、§6 ルール 6）、同じ範囲で既に宣言されている名前、変更後に別の宣言を隠す・別の宣言に隠される名前（モジュールで Excel / VBA のグローバルとして使っている名前、使用箇所の関数のローカル変数と同じ名前など）。拒否した場合は理由を表示し、何も書き換えない。

#### 16.2.6 セマンティックトークン

* トークンの種類は LSP 標準の `namespace` / `type` / `class` / `enum` / `interface` / `struct` / `typeParameter` / `parameter` / `variable` / `property` / `enumMember` / `function` / `method`。修飾子は標準の `declaration` / `readonly` / `static` / `defaultLibrary`（同梱の型定義の名前：Excel / VBA のメンバーとグローバル）/ `deprecated`（JSDoc に `@deprecated` のある宣言）と、RAS 独自の `byref`（参照渡しの引数）/ `extension`（拡張メソッド）。
* `.` で呼んだ拡張メソッドは `method` に `extension` を付ける。VS Code 拡張は `byref` を下線、`deprecated` を取り消し線で表示する既定の設定を持つ。

#### 16.2.7 未インポートの候補

* 候補にするのは、プロジェクトの他のモジュール（フォームのコードビハインドを除く）の `export` で、そのモジュールのスコープにまだ無い名前。拡張メソッドは `.` の後で、レシーバの型に適用できるものだけを候補にする。
* 確定時の `import` は、同じモジュールからの `import { … }` があればその名前の並びの末尾に、無ければ最後の `import` の次の行（`import` が無ければファイルの先頭）に、相対パス（`./util`、`../lib/str`）で追加する。
* 同じ規則で、`名前が見つかりません`（関数・変数・クラス・列挙型）、`型が見つかりません`、拡張メソッドとしてのメンバーが見つからないエラーに、`import` を追加するクイックフィックスを出す。

---

## 17. 参考実装・採用ライブラリ

選定方針: 単一の静的バイナリにまとめられること、ライセンスが MIT / Apache-2.0 等の寛容なものであること、メンテナンスが継続していること。GPL の実装は設計の参考にとどめ、コードは流用しない。各ライブラリのバージョン・ライセンスは導入時に確認する。

### 17.1 コンパイラ本体（Rust）

| 用途 | 採用 | 理由・備考 |
|---|---|---|
| 字句解析 | `logos` | 高速な字句解析器の生成 |
| 構文木 | `rowan` | rust-analyzer のロスレス CST。空白・コメントを保持するため、フォーマッタ・LSP・JSDoc の取得に向く |
| インクリメンタル計算 | `salsa` | rust-analyzer で使われるクエリベースの増分計算。LSP の応答性能（§16.2.4）の基盤 |
| 構文解析 | 手書きの再帰下降 | 参考: `oxc_parser`（Rust 製の高速な TS パーサー。TS 文法の扱い）、Biome の JS/TS パーサー（`rowan` 系 CST 上でのエラー回復）、rust-analyzer のパーサー（イベント方式でのツリー構築）。RAS は `byref` / `inline` 等の独自構文を持つため、既存の TS パーサーの流用はしない |
| 位置情報の変換 | `line-index` | rust-analyzer 由来。LSP の UTF-16 位置とのバイト位置の相互変換 |
| JSDoc の解析 | 自作 | 参考: oxc の JSDoc 解析の実装 |
| フォーマッタ | 自作（中間表現＋プリンタ方式） | 参考: Biome の `biome_formatter`、Prettier の設計 |
| CLI のエラー表示 | `annotate-snippets` | rustc と同じ形式のソース付きエラー表示 |
| CLI 引数 | `clap` | |
| 設定ファイル | `toml`, `serde` | `ras.toml` |
| ファイル監視 | `notify`, `notify-debouncer-full` | `ras watch` |
| 並列化 | `rayon` | モジュール単位の並列コード生成 |
| コードページ変換 | `encoding_rs` | CP932（Windows-31J 相当の `Shift_JIS`）等への変換と、表現できない文字の検出（§13.7） |
| CRC32 | `crc32fast` | Bridge 側の `RtlComputeCrc32` と同じ CRC-32 で、内容ハッシュを照合する（§14.6） |
| テスト | `insta`（スナップショット）、`proptest`、`cargo-fuzz` | RAS → VBA 出力のスナップショットテスト、パーサーのファジング |

### 17.2 LSP

| 用途 | 採用 | 理由・備考 |
|---|---|---|
| LSP の通信層 | `lsp-server` + `lsp-types` | rust-analyzer の同期型の LSP 基盤。`salsa` と組み合わせたスレッド設計を自前で制御できる |
| 代替案 | `tower-lsp-server` | 非同期（tokio）の高レベルフレームワーク（元の `tower-lsp` のコミュニティフォーク）。`lsp-server` で実装負担が大きすぎた場合の代替 |

### 17.3 オフライン書き込み（§14.9）

Rust で `vbaProject.bin` を**書き込める**ライブラリは見当たらない（`ovba` は読み取り専用）ため、書き込み側は自作する。

| 用途 | 採用・参考 | 備考 |
|---|---|---|
| OPC（zip）の読み書き | `zip` | 変更しないエントリは再圧縮せず raw コピーする |
| 複合ファイル（CFB）の読み書き | `cfb` | `vbaProject.bin` のストレージ・ストリーム操作 |
| MS-OVBA 圧縮・展開 | 自作 | 参考: Kavod.Vba.Compression（C#, MIT）、NetOfficeFw/VbaCompression（C#）、Beakerboy/MS-OVBA-Compression（Python） |
| `vbaProject.bin` の生成手順 | 自作 | 参考: Beakerboy/vbaProject-Compiler（Python。ソースから `vbaProject.bin` を生成）、NetOfficeFw/vbamc（C#。`vbaProject.bin` を生成するコンパイラ） |
| 読み取り・テストの検証用 | `ovba`（Rust, 読み取り専用） | 書き出した `vbaProject.bin` を別実装で読み直す相互検証に使う（§14.9.3）。`calamine` の VBA 読み取り部も参考 |
| 検証ツール | oletools（`olevba`, Python） | CI で出力ファイルを解析し、ソース・`dir` ストリームの整合を確認する |
| OPC のパート構成 | 参考: XlsxWriter の `add_vba_project` | `[Content_Types].xml` とリレーションの書き方 |
| ユーザーフォーム（§14.11） | xlflow（Go, MIT）の `internal/vba/userforms` を移植 | FormSpec・検査・MS-OFORMS デザイナストレージの生成と読み取り・画像（`StdPicture`）。テストは xlflow の出力との比較（生成はバイト単位、読み取りは FormSpec の JSON） |
| 仕様 | [MS-OVBA], [MS-CFB], [MS-VBAL], [MS-OFORMS] | 一次資料。実装はこれを正とする |

### 17.4 Windows 連携（ライブ同期 §14.1〜14.8、`ras typegen`）

| 用途 | 採用・参考 | 備考 |
|---|---|---|
| Win32 API・COM | `windows` | 共有メモリ（`CreateFileMappingW`）、ROT（`GetRunningObjectTable`）、`IDispatch`、`ITypeLib` / `ITypeInfo`（`ras typegen`） |
| `IDispatch` の呼び出し | 自作の薄いラッパー | `windows` クレートは低レベル API のみのため、名前による呼び出し（`GetIDsOfNames` → `Invoke`）と `VARIANT` 変換を自作する |
| VBE の操作・同期方式 | 参考: Ariawase の `vbac.wsf`（VBA のエクスポート/インポート）、Rubberduck（C#, GPL-3） | Rubberduck は GPL のため設計の参考のみ |

### 17.5 言語設計・静的解析の参考

| 参考 | 用途 |
|---|---|
| [MS-VBAL]（VBA 言語仕様） | 出力 VBA の意味論（評価順序、型変換、既定メンバー等）の一次資料 |
| Rubberduck のインスペクション一覧 | VBA の罠（§13）とリンタールールの網羅性の確認（設計の参考のみ） |
| twinBASIC | VBA 互換言語の拡張（ジェネリクス等）の設計の参考（商用。参考のみ） |
| TypeScript / C# の言語仕様 | 構文（TS）と数値型・拡張メソッド・ジェネリクスの意味論（C#） |

### 17.6 VS Code 拡張

| 用途 | 採用 | 備考 |
|---|---|---|
| LSP クライアント | `vscode-languageclient` | `ras lsp --stdio` を起動するのみ（§7.4） |
| TextMate 文法 | microsoft/TypeScript-TmLanguage を元に派生 | `byref` / `byval` / `inline` / `noinline` / `resume` と数値リテラルのサフィックスを追加 |
| バンドル・パッケージ | `esbuild`, `@vscode/vsce` | |

---

## 18. 未決事項

現時点でなし。

### 決定済み（v0.3）

* 実装言語は Rust（§7.2）。
* Mac 版 Excel は対象外（§1）。
* オフライン書き込みは初版から対応（§14.9）。
* Range 由来の2次元配列は 0 始まり、ゼロコピーで再基底化（§9.2.1）。

### 決定済み（v0.4）

* 数値型は C# 風（`short` / `int` / `long` / `float` / `double`）。`number` は廃止（§2）。
* C# 風の拡張メソッドに対応（§3.2）。

### 決定済み（v0.5）

* 長時間ループには、間引き付きで `DoEvents` を自動挿入する（§10.4）。
* Access / Word は当面サポートしない（§1）。

### 決定済み（v0.6）

* ジェネリクスは、プロジェクト内で使われる型引数ごとに別の関数・クラスとして出力する（§2.4）。

### 決定済み（v0.8）

* データの入れ物は `Type` に変換する構造体を基本とする。オブジェクトリテラル・タプル・分割代入は構造体の構文として提供し、値が構造体の `Map` はコンパイラがクラスに置き換える（§2.5）。

### 決定済み（v0.9）

* Excel / VBA / Office のメンバーとグローバルは camelCase で参照し、生成する VBA にはタイプ ライブラリの綴りを出力する。VBA の綴りも受け付けるが警告とし、自動で書き換えられるようにする（§15.3.1）。
* ホストのコレクションの配列メソッドは `For Each` に展開し、結果は配列で返す（§4.4）。
* ユーザーフォームのコードビハインドは RAS で書き（`forms/<Name>.ras`）、フォームのコードモジュールにコンパイルする。イベントは `@on("<コントロール>.<イベント>")` で受け取り、`this` がフォームを表す。`ras form pull` は VBA のコードを RAS に変換する（§14.11.2、§14.11.6）。
* 同期（ライブ同期・`ras pack`）は、ブックの手書きの標準モジュール・シート / `ThisWorkbook` のコード・VBE で作ったフォームを RAS に変換して取り込み、変換しきれたものから RAS の生成コードで置き換える（§14.7.1）。非管理モジュールを `vba/` で双方向に同期する方式（§14.7）より先に、RAS 管理へ移す方式を実装する。
