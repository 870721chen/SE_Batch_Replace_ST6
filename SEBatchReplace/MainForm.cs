using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace SEBatchReplace
{
    public sealed class MainForm : Form
    {
        private TextBox _folderTextBox;
        private ListBox _assemblyList;
        private ListBox _mappingList;
        private RadioButton _skipRadio;
        private RadioButton _stopRadio;
        private TextBox _logTextBox;
        private Button _revisionReadButton;
        private Button _asmScanButton;
        private Button _backgroundOpenButton;
        private Button _apiProbeButton;
        private Button _singleReplaceTestButton;
        private readonly List<ReplacementMapping> _mappings = new List<ReplacementMapping>();

        public MainForm()
        {
            Text = "SE Batch Replace ST6";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(900, 620);
            Size = new Size(1080, 720);
            Font = new Font("Segoe UI", 9F);
            BuildInterface();

        }

        private void BuildInterface()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 5
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 25F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 38F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 84));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 37F));
            Controls.Add(root);

            var folderPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            folderPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            folderPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130F));
            folderPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 22F));
            folderPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var folderLabel = new Label
            {
                Text = "組立根資料夾（包含所有子資料夾，遞迴搜尋 .asm）",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 0, 0, 2)
            };
            _folderTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 2, 8, 0),
                ReadOnly = true
            };
            var browseButton = new Button
            {
                Text = "選擇資料夾",
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };
            browseButton.Click += BrowseFolder;
            folderPanel.Controls.Add(folderLabel, 0, 0);
            folderPanel.SetColumnSpan(folderLabel, 2);
            folderPanel.Controls.Add(_folderTextBox, 0, 1);
            folderPanel.Controls.Add(browseButton, 1, 1);
            root.Controls.Add(folderPanel, 0, 0);

            var assemblyGroup = new GroupBox
            {
                Text = "找到的組立檔（包含所有子資料夾）",
                Dock = DockStyle.Fill
            };
            _assemblyList = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true };
            assemblyGroup.Controls.Add(_assemblyList);
            root.Controls.Add(assemblyGroup, 0, 1);

            var mappingGroup = new GroupBox
            {
                Text = "手動指定零件對應（舊零件名稱 → 新 .par 檔案）",
                Dock = DockStyle.Fill
            };
            var mappingLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(6)
            };
            mappingLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            mappingLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130F));
            _mappingList = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true };
            var mappingButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false
            };
            var addButton = new Button { Text = "新增配對", Width = 110, Height = 32 };
            addButton.Click += AddMapping;
            var removeButton = new Button { Text = "移除配對", Width = 110, Height = 32 };
            removeButton.Click += RemoveMapping;
            mappingButtons.Controls.Add(addButton);
            mappingButtons.Controls.Add(removeButton);
            mappingLayout.Controls.Add(_mappingList, 0, 0);
            mappingLayout.Controls.Add(mappingButtons, 1, 0);
            mappingGroup.Controls.Add(mappingLayout);

            var policyGroup = new GroupBox
            {
                Text = "遇到關聯失效時的處理方式",
                Dock = DockStyle.Fill
            };
            _skipRadio = new RadioButton
            {
                Text = "回復該次嘗試、跳過並繼續後續替換",
                AutoSize = true,
                Location = new Point(12, 24),
                Checked = true
            };
            _stopRadio = new RadioButton
            {
                Text = "回復該次嘗試、立即停止批次作業",
                AutoSize = true,
                Location = new Point(12, 50)
            };
            policyGroup.Controls.Add(_skipRadio);
            policyGroup.Controls.Add(_stopRadio);

            var configurationPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            configurationPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58F));
            configurationPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42F));
            configurationPanel.Controls.Add(mappingGroup, 0, 0);
            configurationPanel.Controls.Add(policyGroup, 1, 0);
            root.Controls.Add(configurationPanel, 0, 2);

            var testPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(0, 6, 0, 0)
            };
            _revisionReadButton = new Button
            {
                Text = "Revision Manager 讀取測試（全部組立）",
                Width = 260,
                Height = 32
            };
            _revisionReadButton.Click += RevisionManagerReadAllAssemblies;
            _asmScanButton = new Button
            {
                Text = "ASM 參照初篩（不開啟 Solid Edge）",
                Width = 250,
                Height = 32
            };
            _asmScanButton.Click += ScanAssemblyReferences;
            _backgroundOpenButton = new Button
            {
                Text = "背景開啟測試（不顯示視窗）",
                Width = 230,
                Height = 32
            };
            _backgroundOpenButton.Click += BackgroundOpenSelectedAssembly;
            _apiProbeButton = new Button
            {
                Text = "COM API 成員診斷（唯讀）",
                Width = 220,
                Height = 32
            };
            _apiProbeButton.Click += ProbeAssemblyApiMembers;
            _singleReplaceTestButton = new Button
            {
                Text = "單一 occurrence 替換測試（不儲存）",
                Width = 250,
                Height = 32
            };
            _singleReplaceTestButton.Click += SingleOccurrenceReplaceTest;
            var hint = new Label
            {
                Text = "使用測試副本；單一替換測試不儲存。",
                AutoSize = true,
                Padding = new Padding(8, 8, 0, 0)
            };
            testPanel.Controls.Add(_revisionReadButton);
            testPanel.Controls.Add(_asmScanButton);
            testPanel.Controls.Add(_backgroundOpenButton);
            testPanel.Controls.Add(_apiProbeButton);
            testPanel.Controls.Add(_singleReplaceTestButton);
            testPanel.Controls.Add(hint);
            root.Controls.Add(testPanel, 0, 3);

            var logGroup = new GroupBox { Text = "作業紀錄", Dock = DockStyle.Fill };
            _logTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false
            };
            logGroup.Controls.Add(_logTextBox);
            root.Controls.Add(logGroup, 0, 4);

            AppendLog("請選擇組立根資料夾，建立舊零件與新零件配對，再執行目前的唯讀讀取測試。");
        }

        private void BrowseFolder(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "選擇包含 Solid Edge .asm 組立檔的資料夾";
                const string defaultAssemblyFolder = @"\\1f\共用資料\0正式圖檔\0 量產機種";
                if (Directory.Exists(defaultAssemblyFolder))
                    dialog.SelectedPath = defaultAssemblyFolder;

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                _folderTextBox.Text = dialog.SelectedPath;
                LoadAssemblyFolder(dialog.SelectedPath);
            }
        }

        private void LoadAssemblyFolder(string folderPath)
        {
            try
            {
                IList<string> files = FolderScanner.GetAssemblyFiles(folderPath);
                _assemblyList.Items.Clear();
                string rootPath = folderPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                foreach (string file in files)
                {
                    string relativePath = file.Substring(rootPath.Length)
                        .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    _assemblyList.Items.Add(relativePath);
                }
                AppendLog("遞迴掃描完成：包含所有子資料夾，共找到 " + files.Count + " 個 .asm 檔案。");
            }
            catch (Exception ex)
            {
                AppendLog("掃描失敗：" + ex.Message);
                MessageBox.Show(this, ex.Message, "掃描失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BackgroundOpenSelectedAssembly(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(_folderTextBox.Text) || !Directory.Exists(_folderTextBox.Text))
            {
                MessageBox.Show(this, "請先選擇有效的組立根資料夾。", "尚未選擇資料夾", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (_assemblyList.SelectedIndex < 0)
            {
                MessageBox.Show(this, "請先在組立清單選取一個測試用 .asm 檔案。", "尚未選取組立", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string relativePath = Convert.ToString(_assemblyList.SelectedItem);
            string assemblyPath = Path.Combine(_folderTextBox.Text, relativePath);
            if (!File.Exists(assemblyPath))
            {
                MessageBox.Show(this, "找不到選取的組立檔：" + assemblyPath, "檔案不存在", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _backgroundOpenButton.Enabled = false;
            UseWaitCursor = true;
            AppendLog("========== 開始 Solid Edge 背景開啟測試 ==========");
            AppendLog("[背景測試] " + relativePath);
            AppendLog("測試模式：隱藏 Solid Edge 主視窗，使用 JDOCUMENTPROP_NOWINDOW=8 開啟組立，不儲存，讀取後關閉。");

            try
            {
                IList<string> occurrences = SolidEdgeReader.OpenReadAndCloseAssemblyInBackground(assemblyPath);
                AppendLog("[背景開啟成功] " + relativePath + "；第一層 occurrence 數量：" + occurrences.Count);
                foreach (string occurrence in occurrences)
                    AppendLog("  " + occurrence);

                MessageBox.Show(this,
                    "背景開啟測試成功。\r\n組立檔：" + relativePath +
                    "\r\n第一層 occurrence：" + occurrences.Count +
                    "\r\n\r\n文件已關閉且未儲存。請確認測試期間 Solid Edge 主視窗及組立視窗均未顯示。",
                    "背景開啟測試成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendLog("[背景開啟失敗] " + GetExceptionDetails(ex).Replace(Environment.NewLine, " | "));
                MessageBox.Show(this, ex.Message, "背景開啟測試失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
                _backgroundOpenButton.Enabled = true;
                AppendLog("========== Solid Edge 背景開啟測試結束 ==========");
            }
        }


        private void ProbeAssemblyApiMembers(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(_folderTextBox.Text) || !Directory.Exists(_folderTextBox.Text))
            {
                MessageBox.Show(this, "請先選擇有效的組立根資料夾。", "尚未選擇資料夾", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (_assemblyList.SelectedIndex < 0)
            {
                MessageBox.Show(this, "請先在組立清單選取一個測試用 .asm 檔案。", "尚未選取組立", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string relativePath = Convert.ToString(_assemblyList.SelectedItem);
            string assemblyPath = Path.Combine(_folderTextBox.Text, relativePath);
            if (!File.Exists(assemblyPath))
            {
                MessageBox.Show(this, "找不到選取的組立檔：" + assemblyPath, "檔案不存在", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _apiProbeButton.Enabled = false;
            UseWaitCursor = true;
            AppendLog("========== 開始 Solid Edge COM API 成員診斷 ==========");
            AppendLog("[API 診斷] " + relativePath);
            AppendLog("唯讀模式：隱藏 Solid Edge，背景開啟組立；只讀取 COM TypeInfo，不呼叫替換、修改或儲存。");
            try
            {
                IList<string> lines = SolidEdgeReader.ProbeAssemblyApiInBackground(assemblyPath);
                foreach (string line in lines)
                    AppendLog(line);
                MessageBox.Show(this,
                    "COM API 成員診斷完成。\r\n組立：" + relativePath +
                    "\r\n診斷結果已寫入作業紀錄。\r\n本次未修改或儲存任何檔案。",
                    "API 診斷完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendLog("[API 診斷失敗] " + GetExceptionDetails(ex).Replace(Environment.NewLine, " | "));
                MessageBox.Show(this, ex.Message, "API 診斷失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
                _apiProbeButton.Enabled = true;
                AppendLog("========== Solid Edge COM API 成員診斷結束 ==========");
            }
        }

        private void SingleOccurrenceReplaceTest(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(_folderTextBox.Text) || !Directory.Exists(_folderTextBox.Text))
            {
                MessageBox.Show(this, "請先選擇有效的組立根資料夾。", "尚未選擇資料夾", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (_assemblyList.SelectedIndex < 0)
            {
                MessageBox.Show(this, "請先選取已複製的測試用 .asm 檔案。", "尚未選取組立", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (_mappingList.SelectedIndex < 0 || _mappingList.SelectedIndex >= _mappings.Count)
            {
                MessageBox.Show(this, "請先選取一筆舊零件與新零件配對。", "尚未選取配對", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string relativePath = Convert.ToString(_assemblyList.SelectedItem);
            string assemblyPath = Path.Combine(_folderTextBox.Text, relativePath);
            ReplacementMapping mapping = _mappings[_mappingList.SelectedIndex];
            if (!File.Exists(assemblyPath))
            {
                MessageBox.Show(this, "找不到選取的組立檔：" + assemblyPath, "檔案不存在", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!File.Exists(mapping.NewPartPath))
            {
                MessageBox.Show(this, "找不到新零件檔案：" + mapping.NewPartPath, "新零件不存在", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult confirmation = MessageBox.Show(this,
                "即將在 Solid Edge 背景開啟選取的測試組立，並只替換一個符合的 occurrence。\r\n\r\n" +
                "組立：" + relativePath + "\r\n舊零件：" + mapping.OldPartName +
                "\r\n新零件：" + mapping.NewPartPath +
                "\r\n\r\n本次設定 ReplaceAll=false，測試後以 Close(false) 關閉，不儲存。\r\n" +
                "如果符合的 occurrence 不是剛好一個，程式會停止而不執行替換。\r\n\r\n是否開始？",
                "確認單一零件替換測試", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirmation != DialogResult.Yes) return;

            _singleReplaceTestButton.Enabled = false;
            UseWaitCursor = true;
            AppendLog("========== 開始單一 occurrence 替換測試 ==========");
            AppendLog("[測試組立] " + relativePath);
            AppendLog("[替換配對] " + mapping.OldPartName + " → " + mapping.NewPartPath);
            AppendLog("[安全設定] ReplaceAll=false；測試後 Close(false)，不儲存。");
            try
            {
                IList<string> lines = SolidEdgeReader.ReplaceOneOccurrenceInBackground(
                    assemblyPath, mapping.OldPartName, mapping.NewPartPath);
                foreach (string line in lines) AppendLog(line);
                MessageBox.Show(this,
                    "單一 occurrence 替換測試已完成。\r\n組立已關閉且未儲存。\r\n請查看作業紀錄中的替換前後路徑及關聯數量。",
                    "替換測試完成（未儲存）", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendLog("[單一替換測試失敗] " + GetExceptionDetails(ex).Replace(Environment.NewLine, " | "));
                MessageBox.Show(this, ex.Message, "單一替換測試失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
                _singleReplaceTestButton.Enabled = true;
                AppendLog("========== 單一 occurrence 替換測試結束 ==========");
            }
        }

        private void ScanAssemblyReferences(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(_folderTextBox.Text) || !Directory.Exists(_folderTextBox.Text))
            {
                MessageBox.Show(this, "請先選擇有效的組立根資料夾。", "尚未選擇資料夾", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (_mappings.Count == 0)
            {
                MessageBox.Show(this, "請先新增至少一筆舊零件名稱配對，才能進行參照初篩。", "尚未設定舊零件", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IList<string> files;
            try
            {
                files = FolderScanner.GetAssemblyFiles(_folderTextBox.Text);
            }
            catch (Exception ex)
            {
                AppendLog("ASM 參照初篩：資料夾掃描失敗：" + ex.Message);
                return;
            }

            int hitFiles = 0;
            int unreadableFiles = 0;
            _asmScanButton.Enabled = false;
            _revisionReadButton.Enabled = false;
            UseWaitCursor = true;
            AppendLog("========== 開始 ASM 參照二進位初篩 ==========");
            AppendLog("組立檔數量：" + files.Count + "；待比對舊零件數量：" + _mappings.Count + "。");
            AppendLog("只讀取檔案位元組，不啟動 Solid Edge、不修改或儲存檔案。");

            try
            {
                for (int i = 0; i < files.Count; i++)
                {
                    string file = files[i];
                    string relative = file.Substring(_folderTextBox.Text.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar).Length)
                        .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    try
                    {
                        AsmReferenceScanResult result = AsmReferenceScanner.Scan(file, GetOldPartNames());
                        AppendLog("[檔案] " + relative + "；格式辨識：" + result.FileFormat + "；檔案大小：" + result.FileLength + " bytes");
                        if (result.Matches.Count > 0)
                        {
                            hitFiles++;
                            AppendLog("[命中] " + relative);
                            foreach (AsmReferenceMatch match in result.Matches)
                                AppendLog("  [參照文字] " + match.PartName + "；編碼：" + match.EncodingName +
                                    "；位元組位置：" + String.Join(", ", match.Offsets) +
                                    (match.Truncated ? "；位置僅列前 10 筆" : ""));
                        }
                    }
                    catch (Exception ex)
                    {
                        unreadableFiles++;
                        AppendLog("[掃描失敗] " + relative + "；" + ex.Message);
                    }
                }
            }
            finally
            {
                UseWaitCursor = false;
                _asmScanButton.Enabled = true;
                _revisionReadButton.Enabled = true;
            }

            AppendLog("========== ASM 參照二進位初篩結束 ==========");
            AppendLog("組立檔：" + files.Count + "；文字命中組立：" + hitFiles + "；讀取失敗：" + unreadableFiles + "。");
            AppendLog("注意：命中代表檔案原始位元組中找到檔名文字，不等於 occurrence 數；未命中也不能證明組立沒有參照。");
            MessageBox.Show(this, "ASM 參照初篩完成。\r\n組立檔：" + files.Count +
                "\r\n文字命中組立：" + hitFiles + "\r\n讀取失敗：" + unreadableFiles +
                "\r\n\r\n詳細結果請查看作業紀錄。\r\n本次沒有修改任何檔案。",
                "ASM 參照初篩完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private IList<string> GetOldPartNames()
        {
            var names = new List<string>();
            foreach (ReplacementMapping mapping in _mappings)
                names.Add(mapping.OldPartName);
            return names;
        }

        private void RevisionManagerReadAllAssemblies(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(_folderTextBox.Text) || !Directory.Exists(_folderTextBox.Text))
            {
                MessageBox.Show(this, "請先選擇有效的組立根資料夾。", "尚未選擇資料夾", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IList<string> assemblyFiles;
            try
            {
                assemblyFiles = FolderScanner.GetAssemblyFiles(_folderTextBox.Text);
            }
            catch (Exception ex)
            {
                AppendLog("重新掃描組立資料夾失敗：" + ex.Message);
                MessageBox.Show(this, ex.Message, "掃描失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _assemblyList.Items.Clear();
            string rootPath = _folderTextBox.Text.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            foreach (string file in assemblyFiles)
            {
                string relativePath = file.Substring(rootPath.Length)
                    .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                _assemblyList.Items.Add(relativePath);
            }

            if (assemblyFiles.Count == 0)
            {
                AppendLog("指定資料夾及其子資料夾中沒有找到 .asm 檔案。");
                MessageBox.Show(this, "指定資料夾及其子資料夾中沒有找到 .asm 檔案。", "沒有組立檔", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int successCount = 0;
            int failedCount = 0;
            int totalLinks = 0;

            AppendLog("========== 開始 Revision Manager 全資料夾讀取 ==========");
            AppendLog("掃描到 " + assemblyFiles.Count + " 個 .asm 檔案；逐一讀取，不依賴清單選取項目。");
            AppendLog("本次只讀取連結資訊，不呼叫 Save、Replace 或任何連結修改方法。");

            _revisionReadButton.Enabled = false;
            UseWaitCursor = true;
            try
            {
                for (int i = 0; i < assemblyFiles.Count; i++)
                {
                    string assemblyPath = assemblyFiles[i];
                    string relativePath = assemblyPath.Substring(rootPath.Length)
                        .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

                    AppendLog("[" + (i + 1) + "/" + assemblyFiles.Count + "] 開始讀取：" + relativePath);
                    try
                    {
                        IList<string> links = RevisionManagerReader.ReadLinkedDocuments(assemblyPath);
                        successCount++;
                        totalLinks += Math.Max(0, links.Count - 1);
                        AppendLog("[讀取成功] " + relativePath + "；回傳資料列：" + links.Count);
                        foreach (string link in links)
                            AppendLog("  [連結] " + link);
                    }
                    catch (Exception ex)
                    {
                        failedCount++;
                        AppendLog("[讀取失敗] " + relativePath + "；" +
                            GetExceptionDetails(ex).Replace(Environment.NewLine, " | "));
                    }
                }
            }
            finally
            {
                UseWaitCursor = false;
                _revisionReadButton.Enabled = true;
            }

            AppendLog("========== Revision Manager 全資料夾讀取結束 ==========");
            AppendLog("組立總數：" + assemblyFiles.Count +
                "；成功：" + successCount +
                "；失敗：" + failedCount +
                "；連結資料列總數（不含各檔案標題列）：" + totalLinks + "。");

            MessageBox.Show(this,
                "Revision Manager 全資料夾讀取完成。" +
                "\r\n組立檔總數：" + assemblyFiles.Count +
                "\r\n成功：" + successCount +
                "\r\n失敗：" + failedCount +
                "\r\n連結資料列：" + totalLinks +
                "\r\n\r\n各組立詳細結果請查看作業紀錄。\r\n本次沒有修改或儲存任何檔案。",
                "Revision Manager 讀取完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static string GetExceptionDetails(Exception exception)
        {
            var details = new System.Text.StringBuilder();
            Exception current = exception;
            int level = 0;
            while (current != null)
            {
                if (level > 0)
                    details.AppendLine();
                details.Append("例外 ").Append(level + 1).Append("：")
                    .Append(current.GetType().FullName).AppendLine();
                details.Append("訊息：").AppendLine(current.Message);
                // Exception.HResult is not publicly readable on older .NET Framework versions.
                // Keep diagnostics compatible with the legacy Windows 7 runtime.
                current = current.InnerException;
                level++;
            }
            return details.ToString();
        }

        private void AddMapping(object sender, EventArgs e)
        {
            string oldPartName = PromptOldPartName();
            if (oldPartName == null)
                return;

            string newPath = SelectNewPartFile();
            if (newPath == null)
                return;

            foreach (ReplacementMapping mapping in _mappings)
            {
                if (String.Equals(mapping.OldPartName, oldPartName, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(this, "這個舊零件名稱已經建立配對。", "重複配對", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            var newMapping = new ReplacementMapping(oldPartName, newPath);
            _mappings.Add(newMapping);
            _mappingList.Items.Add(newMapping);
            AppendLog("新增配對：" + newMapping.OldPartName + " → " + Path.GetFileName(newPath));
        }

        private string PromptOldPartName()
        {
            using (var dialog = new Form())
            using (var nameBox = new TextBox())
            using (var label = new Label())
            using (var okButton = new Button())
            using (var cancelButton = new Button())
            {
                dialog.Text = "輸入舊零件名稱";
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;
                dialog.ShowInTaskbar = false;
                dialog.ClientSize = new Size(390, 112);

                label.Text = "請輸入組立檔中原有的零件檔名（可省略 .par）：";
                label.AutoSize = true;
                label.Location = new Point(12, 12);
                nameBox.Location = new Point(12, 38);
                nameBox.Width = 365;
                okButton.Text = "確定";
                okButton.DialogResult = DialogResult.OK;
                okButton.Location = new Point(210, 73);
                cancelButton.Text = "取消";
                cancelButton.DialogResult = DialogResult.Cancel;
                cancelButton.Location = new Point(292, 73);

                dialog.Controls.Add(label);
                dialog.Controls.Add(nameBox);
                dialog.Controls.Add(okButton);
                dialog.Controls.Add(cancelButton);
                dialog.AcceptButton = okButton;
                dialog.CancelButton = cancelButton;

                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return null;

                string value = nameBox.Text.Trim();
                if (String.IsNullOrWhiteSpace(value))
                {
                    MessageBox.Show(this, "請輸入舊零件名稱。", "名稱不可空白", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }

                if (!String.Equals(Path.GetFileName(value), value, StringComparison.Ordinal))
                {
                    MessageBox.Show(this, "請只輸入檔名，不需要輸入資料夾路徑。", "名稱格式錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }

                if (value.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                {
                    MessageBox.Show(this, "零件名稱包含無效字元。", "名稱格式錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }

                if (value.IndexOf('.') >= 0 && !value.EndsWith(".par", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(this, "請輸入 .par 零件檔名。", "名稱格式錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }

                if (!value.EndsWith(".par", StringComparison.OrdinalIgnoreCase))
                    value += ".par";

                return value;
            }
        }

        private string SelectNewPartFile()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "選擇要替換成的新零件（.par）";
                dialog.Filter = "Solid Edge 零件檔 (*.par)|*.par";
                dialog.CheckFileExists = true;
                return dialog.ShowDialog(this) == DialogResult.OK ? dialog.FileName : null;
            }
        }

        private void RemoveMapping(object sender, EventArgs e)
        {
            int index = _mappingList.SelectedIndex;
            if (index < 0) return;
            _mappingList.Items.RemoveAt(index);
            _mappings.RemoveAt(index);
            AppendLog("已移除選取的零件配對。");
        }

        private void AppendLog(string message)
        {
            _logTextBox.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + Environment.NewLine);
        }
    }
}
