# 駅すぱあと API SDK for VBA

[駅すぱあと API](https://docs.ekispert.com/v1/index.html)をExcel VBAなどから利用するためのSDKです。

## 使い方

[バイナリファイルをダウンロード](https://github.com/EkispertAPIMania/VBA-SDK/releases/)するか、vbacを使ってビルドします。

```
build.bat
```

マクロを開いて、実装します。

## 初期化

初期化時には、駅すぱあと APIのAPIキーを指定します。[APIキーはトライアル申し込みより取得](https://api-info.ekispert.com/form/trial/)してください。

```vb
Dim client As Ekispert
Set client = New Ekispert
client.ApiKey = "YOUR_API_KEY"
```

## 駅情報の取得

駅情報取得APIを実行します。検索条件、結果は[駅情報 - 駅すぱあと API Documents 駅データ・経路検索のWebAPI](https://docs.ekispert.com/v1/api/station.html)を参照してください。

```vb
Dim Query As StationQuery
Set Query = client.StationQuery()
Query.Name = "東京" ' 駅名で検索
Dim Result As ResultSet
Result = Query.Find()

Debug.Print Result.Max ' 200（検索結果数）
Debug.Print Result.Points(0).Station.Code ' 22828
Debug.Print Result.Points(0).GeoPoint.Latitude_DD ' 35.678083
Debug.Print Result.Points(0).GeoPoint.Longitude_DD ' 139.770444
Dim i As Long

For i = 0 To UBound(Result.Points)
		Debug.Print i & " " & Result.Points(i).Station.Code
Next i
```

## 駅情報の取得（簡易）

無料で使える駅簡易情報取得を使う場合です。検索条件、結果は[駅簡易情報 - 駅すぱあと API Documents 駅データ・経路検索のWebAPI](https://docs.ekispert.com/v1/api/station/light.html)を参照してください。

```vb
Dim Query As StationLightQuery ' 駅簡易情報取得用クエリー
Set Query = client.StationLightQuery()
Query.Name = "東京" ' 駅名で検索
Dim Result As ResultSet
Result = Query.Find()

Dim i As Long

For i = 0 To UBound(Result.Points)
		Debug.Print i & " " & Result.Points(i).Station.Name
		Debug.Print i & " " & Result.Points(i).Station.Yomi
Next i
```

## 依存ライブラリ

すべてMITライセンスのライブラリを使用しています。

- [VBA-tools/VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)
- [VBA-tools/VBA-Web: VBA-Web](https://github.com/VBA-tools/VBA-Web)
- [VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON)

## ライセンス

MITライセンスです。
