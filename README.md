# SE Batch Replace ST6

Solid Edge ST6 批次零件替換工具（Windows 桌面程式）。

## 目前階段

目前程式暫時收斂為 **Revision Manager 全資料夾唯讀讀取測試**。使用者選擇組立根資料夾後，程式會遞迴掃描根目錄及所有子資料夾中的 .asm，再逐一透過 Revision Manager 嘗試讀取每個組立的 LinkedDocuments 直接連結清單。

單一組立讀取失敗時會記錄錯誤並繼續下一個組立。測試只讀取連結資訊，不呼叫 Save、Replace 或其他連結修改方法，不會修改或儲存組立與零件檔。

目前介面保留資料夾選擇、組立檔清單、Revision Manager 全資料夾讀取測試及作業紀錄。其他舊測試、替換預覽、零件配對與批次替換介面暫時移除，待讀取方案驗證後再逐步恢復需要的功能。

## 目標環境

- Windows 7 SP1 x64
- Solid Edge ST6
- C# / Windows Forms
- .NET Framework 4.8（Windows 7 SP1 可安裝；目標電腦需安裝 .NET Framework 4.8）
- x64
- GitHub Actions：Windows Server 2022 x64

> 目前以 COM late binding 為方向，避免 GitHub Runner 因未安裝 Solid Edge ST6 而無法編譯。真正連線與操作仍須在安裝 ST6 的電腦測試。

## Revision Manager 讀取測試（實驗功能）

1. 選擇包含組立檔的根資料夾。
2. 程式遞迴列出所有子資料夾中的 .asm。
3. 點選「Revision Manager 讀取測試（全部組立）」；不需要選取清單中的特定組立。
4. 程式逐一讀取所有組立的 LinkedDocuments，並記錄成功、失敗及回傳連結資料。
5. 單一檔案失敗不會中止整批讀取。

**重要限制：** 此功能仍是實驗性測試。Revision Manager 的 COM ProgID、Open、LinkedDocuments、Item、FullName 等成員尚未完整確認其在 Solid Edge ST6 環境中的行為。直接連結清單也不等於組立 occurrence 實例清單；同一零件的多個 instance 可能只呈現一筆連結，且目前尚未遞迴展開子組立。測試結果必須與已知組立結構比對，才能判斷完整性。

請先使用測試資料副本。程式不會呼叫 Save、Replace 或其他連結修改方法，也不會主動修改或儲存任何檔案。

## 編譯

GitHub Actions 需由使用者手動觸發編譯：

1. 開啟 Repository 的 **Actions**。
2. 選擇 **Build SE Batch Replace**。
3. 等待工作完成。
4. 在成功的執行紀錄底部下載 SEBatchReplace-x64 Artifact。

## 本機測試需求

在 Windows 7 SP1 x64 電腦上：

- 安裝 Solid Edge ST6。
- 安裝 .NET Framework 4.8。
- 使用 x64 版本執行程式。
- 先以測試用組立檔與零件副本驗證，不要直接對正式資料執行。

## 版本狀態

目前只保留 Revision Manager 全資料夾唯讀讀取測試介面。真正的 Solid Edge 批次零件替換功能尚未實作；Revision Manager 讀取結果也尚未完成 ST6 實機驗證。
