# StackTraceLib_VBA
これは Microsoft Excel VBA においてコードのデバッグを補助する目的で書くプロシージャの呼出関係を記録しワークシートへ書き出すためのライブラリです。

## Usage - 使用方法
- デバッグの対象とする Excel VBA のカレントプロジェクトへ `src/StackTrace.bas` と `src/StackTraceLog.cls` をインポートします。
- `bin/Trace.xlsm` を開き `DebugTrace` という名称のワークシートをデバッグの対象とする Excel ブックへコピーします。
- ワークシート `DebugTrace` 内に配置された `AddStackTrace` と書かれたボタンを押すと `src/StackTrace.bas` の中にある `AddStackTrace` サブルーチンが実行され Excel VBA のカレントプロジェクトの中にあるすべてのプロシージャへ `PushStackTrace` と `PopStackTrace` の呼出が追加されます。
- ワークシート `DebugTrace` 内に配置された `EnableDebugMode` と書かれたボタンを押すと `src/StackTrace.bas` の中にある `EnableDEBUGMODE` サブルーチンが実行され Excel VBA のカレントプロジェクトの中にあるすべてのモジュールへ `#Const DEBUG_MODE = 1` が追記され `PushStackTrace` と `PopStackTrace` の呼出が有効化されます。
- この状態でデバッグの対象とする Excel VBA のコードを実行すると `src/StackTrace.bas` の中にある `DebugTrace` コレクションへプロシージャの呼出関係が記録されます。
- コードの実行後にワークシート `DebugTrace` 内に配置された `WriteStackTrace Here` と書かれたボタンを押すと `src/StackTrace.bas` の中にある `WriteStackTrace` サブルーチンが実行され `DebugTrace` コレクションへ記録されたプロシージャの呼出関係がワークシート `DebugTrace` へ出力されます。

## Option - オプション

- `src/StackTrace.bas` モジュール内で `#Const DEBUG_PRINT_MODE = 1` を宣言すると `WriteStackeTrace` を実行しなくてもイミディエイトウィンドウへ途中経過が出力されます。（最後まで走らないプログラムをデバッグしたい時に使用します。）
- 各モジュールのコードの1～9行目へ `#Const NO_TRACE = 1` を宣言すると，そのモジュールは `AddStackTrace` の対象から除外することができます。
- 同一プロシージャからの `PushStackTrace` の呼び出し回数が1万回を超えると，デバッグを中断し，メッセージボックスが表示されます。（応答なしとなることを防ぐための措置です。）
  - そのように大量に呼出されるプロシージャは，値を取得するためのProperty Get関数や比較関数などが想定されます。このようなプロシージャは単純処理を行うものが多く，トレースの対象から除外してもコードの理解にはあまり影響がないことが多いため，このメッセージが表示されたプロシージャからは手作業で `PushStackTrace` と `PopStackTrace` の呼出コードを削除することが推奨されます。
  - なお，このメッセージが表示された後，デバッグを終了せずに，そのままデバッグを再開すると，以降は処理を中断せずに実行します。コードの修正をする場合は，ここでデバッグを終了させてください。

## Description - 詳細

- `StackTrace.AddStackTrace` を実行すると，VBAProject内の全てのモジュールの全てのプロシージャが，はじめに `StackTrace.PushStackTrace` を呼出し，最後に `StackTrace.PopStackTrace` を呼出すように，コードに変更が加えられます。
  - この変更は `StackTrace.RemoveStackTrace` により取り除くことができます。
  - `PushStackTrace` の呼出コードは各プロシージャの宣言部の次行へ挿入されます。
  - `PopStackTrace` の呼出コードは各プロシージャの最終行，Exit Sub/Functionの前行，1行形式IF文で記述されたExit Sub/Functionの場合はコロン区切りの１行形式でExit文の直前へ挿入されます。
  - 1行形式のプロシージャ定義には挿入されません。
  - 各モジュールのコードの1～9行目へ `#Const NO_TRACE = 1` を宣言した場合も挿入されません。
- 各プロシージャから `StackTrace.PushStackTrace` が呼出されると，スタックのレベル(`StackTrace.StackLevel`)を1増やすとともに，その時のレベル，モジュール名，プロシージャ名，引数のリストを文字列化したものを `StackTraceLog` オブジェクトとして `StackTrace.DebugTrace` コレクションへ記録します。
- 各プロシージャから `StackTrace.PopStackTrace` が呼出されると，スタックのレベル(`StackTrace.StackLevel`)を1減らすとともに，呼出元プロシージャが `StackTrace.PushStackTrace` した時に追加された `StackTraceLog` オブジェクトへ戻り値を記録します。
- この時，引数リストや戻り値の内容は，`CStr`関数により文字列化されます。`CStr`関数で文字列化できない変数（`Attribute Value.VB_UserMemId = 0` の設定されたデフォルトプロパティを持たないクラスのオブジェクト，配列等）は，`TypeName`の文字列へ角括弧を加えた文字列となります。なお，配列の場合は要素数を括弧内に含めた文字列となります。（`StackTrace.ArgsToString`関数および`StackTrace.PrintArrayBounds`関数参照）
- プロシージャの宣言が `Property Get` の場合は，引数リストの部分に `[Property-Get]`と表示し，戻り値は `関数名＝戻り値` の形式で表示します。
