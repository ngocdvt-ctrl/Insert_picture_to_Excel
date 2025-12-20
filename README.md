# Insert_picture_to_Excel
Autmatically insert .jpg pictures to Excel without spending time → Reduce report making time
![image URL](https://github.com/ngocdvt-ctrl/Insert_picture_to_Excel/blob/main/demo.jpg?raw=true)

Excel画像自動挿入ツール (VBA)
このプロジェクトは、Excelのセル値に基づいて、指定されたフォルダから画像を自動的に挿入するVBAマクロを提供します。商品カタログの作成、画像付きレポート、在庫管理などに最適なソリューションです。

🌟 主な機能
自動挿入: セルに記入されたファイル名に基づき、データ範囲全体をスキャンして画像を挿入します。
自動リサイズ: 挿入された画像は自動的に中央に配置され、セルのサイズに合わせて調整されます（デフォルトはセルサイズの98%）。
スマート処理: 重複を防ぐため、新しい画像を挿入する前にセル内の古い画像を自動的に削除します。
動的配置: xlMoveAndSize プロパティを使用しているため、行や列のサイズを変更しても画像が追従します。

📁 フォルダ構成
コードを正常に動作させるために、以下のようにファイルを配置してください：
3_Insert_picture_to_Excel
├── InsertVBA.xlsm         # VBAマクロが含まれるファイル（標準コード）
├── target.xlsx         # 画像を挿入する対象のExcelファイル
└── [image_files].jpg   # サンプル画像ファイル (1.jpg, 2.jpg, 3.jpg)

🛠 使い方
このリポジトリをローカルPCにダウンロードします。
画像ファイルとExcelファイルが同じフォルダにあることを確認してください。
InsertMacro.xlsm を開きます。
INSERTボタンを押します。
target.xlsx の Sheet1 で結果を確認します。

📝 標準コードの設定について
コード内の定数を変更することで、簡単にカスタマイズが可能です：
target.xlsx: 画像を挿入したいExcelファイルの名前。
IMAGE_FILE_EXTENSION: 画像の拡張子（デフォルトは .jpg）。
SIZE_FACTOR: セルに対する画像の倍率（0.98 = 98%）。
