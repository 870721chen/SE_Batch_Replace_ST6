# SE Batch Replace ST6

Solid Edge ST6 批次零件替換工具（Windows 桌面程式）。

## 目前階段

此 Repository 先建立可由 GitHub Actions 編譯的 Windows Forms 專案骨架。現階段包含資料夾掃描、組立檔清單、手動新舊零件配對介面與錯誤處理策略選擇；**尚未執行 Solid Edge 零件替換**。

Solid Edge COM 操作會在確認 ST6 安裝環境與實際 API 後再實作，避免臆測 API 或宣稱尚未驗證的關聯保留／復原能力。

## 目標環境

- Windows 7 SP1 x64
- Solid Edge ST6
- C# / Windows Forms
- .NET Framework 4.8（Windows 7 SP1 可安裝；目標電腦需安裝 .NET Framework 4.8）
- x64
- GitHub Actions：Windows Server 2022 x64

> 目前以 COM late binding 為方向，避免 GitHub Runner 因未安裝 Solid Edge ST6 而無法編譯。真正連線與操作仍須在安裝 ST6 的電腦測試。

## 使用者已確認的操作規格

1. 使用者選擇一個資料夾。
2. 只掃描該資料夾第一層的 `.asm`，不遞迴子資料夾。
3. 使用者手動指定舊 `.par` 與新 `.par` 的對應。
4. 替換必須盡可能保留原有組立關係。
5. 關聯失效時，執行前由使用者選擇：
   - 跳過該次替換並繼續。
   - 回復該次嘗試並立即停止。
6. 記錄每個組立檔、零件、結果與原因。

## 安全原則

- 不會在未驗證 ST6 API 前假設替換可原子化或可可靠回復。
- 正式批次作業應優先評估工作副本／暫存工作區，避免直接破壞來源檔。
- 任何無法確認的關聯狀態都不得回報為成功。
- 開發階段不會自動覆寫使用者原始組立檔。

## 編譯

GitHub Actions 會在推送至 `main` 或手動觸發時編譯：

1. 開啟 Repository 的 **Actions**。
2. 選擇 **Build SE Batch Replace**。
3. 等待工作完成。
4. 在成功的執行紀錄底部下載 `SEBatchReplace-x64` Artifact。

## 本機測試需求

在 Windows 7 SP1 x64 電腦上：
- 安裝 Solid Edge ST6。
- 安裝 .NET Framework 4.8。
- 使用 x64 版本執行程式。
- 先以測試用組立檔與零件副本驗證，不要直接對正式資料執行。

## 版本狀態

目前為初始開發骨架，尚未具備真正的 Solid Edge 批次替換功能。
