# set password to excel with dotnet

like poi(https://github.com/apache/poi)

## dependency
- OpenMcdf https://github.com/ironfede/openmcdf
  - License MPL-2.0 https://github.com/ironfede/openmcdf?tab=MPL-2.0-1-ov-file#readme 

## run

at p2 dir, run, and check b.xlsx(hard coding now)

```
ers/seijiro/Sync/sync_work/me/SetPassToExceldotNet/p2/bin/Debug/net9.0/p2
暗号化完了: /Users/seijiro/Sync/sync_work/me/SetPassToExceldotNet/b.xlsx

=== 復号化テスト ===

dotnet版 を復号化中...
  復号化後サイズ: 6819 bytes
  最初の16バイト: 50 4B 03 04 14 00 08 08 08 00 C7 6B 43 5B 00 00
  ✓ 正常なZIPファイル（PKシグネチャ確認）

poi版 を復号化中...
  復号化後サイズ: 6819 bytes
  最初の16バイト: 50 4B 03 04 14 00 08 08 08 00 C7 6B 43 5B 00 00
  ✓ 正常なZIPファイル（PKシグネチャ確認）

=== 元ファイルとの比較 ===
元ファイル: 6819 bytes
dotnet復号化: 6819 bytes
poi復号化: 6819 bytes

dotnet版と元ファイル: ✓ 完全一致
poi版と元ファイル: ✓ 完全一致

処理は終了コード 0 で終了しました。
```
