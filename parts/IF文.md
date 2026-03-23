# IF文テンプレート

## フォーマット（値比較）

```
『IF文』
「{column}」が""{value}""《{condition}》の場合、""{true_value}""に変換
いずれの条件分岐にも該当しない場合、{else_value}に変換
```

## フォーマット（カラム値比較）

```
『IF文』
「{column_a}」が「{column_b}」《次のカラムの値に完全一致》の場合、""{true_value}""に変換
いずれの条件分岐にも該当しない場合、《空白》に変換
```

## フォーマット（複数条件）

```
『IF文』
「{column}」が""{value_1}""《{condition}》の場合、""{result_1}""に変換
「{column}」が""{value_2}""《{condition}》の場合、""{result_2}""に変換
いずれの条件分岐にも該当しない場合、""{else_value}""に変換
```

## else_value
- 《空白》
- ""{具体的な値}""
- 《元の値のまま》
