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
        private ComboBox _readScopeComboBox;
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
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
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
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 0)
            };
            var revisionReadButton = new Button
            {
                Text = "Revision Manager 讀取測試（全部組立）",
                Width = 260,
                Height = 32
            };
            revisionReadButton.Click += RevisionManagerReadAllAssemblies;
            var hint = new Label
            {
                Text = "逐一讀取資料夾內所有 .asm；唯讀，不修改、不儲存檔案。",
                AutoSize = true,
                Padding = new Padding(8, 8, 0, 0)
            };
            testPanel.Controls.Add(revisionReadButton);
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
                if (dialog.ShowDialog(this) != DialogResult.OK)
                    return;

                _folderTextBox.Text = dialog.SelectedPath;
                try
                {
                    IList<string> files = FolderScanner.GetAssemblyFiles(dialog.SelectedPath);
                    _assemblyList.Items.Clear();
                    string rootPath = dialog.SelectedPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
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
        }

        private void ReadActiveAssembly(object sender, EventArgs e)
        {
            try
            {
                AppendLog("開始連接 Solid Edge，讀取目前作用中的組立結構……");
                bool includeNested = _readScopeComboBox.SelectedIndex == 1;
                string scopeName = includeNested ? "完整階層" : "主組立第一層";
                IList<string> occurrences = SolidEdgeReader.ReadActiveAssemblyOccurrences(includeNested);
                AppendLog("讀取完成：" + scopeName + "共取得 " + occurrences.Count + " 個零件／子組立參照。");
                foreach (string item in occurrences)
                    AppendLog(item);

                MessageBox.Show(this,
                    "讀取完成（" + scopeName + "），共取得 " + occurrences.Count + " 個參照。\r\n詳細資料已列在作業紀錄。",
                    "組立結構讀取完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string details = GetExceptionDetails(ex);
                AppendLog("讀取失敗： " + details);
                MessageBox.Show(this, details, "Solid Edge 讀取失敗（診斷資訊）", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void AutoReadSelectedAssembly(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(_folderTextBox.Text))
            {
                MessageBox.Show(this, "請先選擇組立根資料夾。", "尚未選擇資料夾", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_assemblyList.SelectedIndex < 0)
            {
                MessageBox.Show(this, "請先在上方組立清單選擇一個 .asm 檔案。", "尚未選擇組立", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string rootPath = _folderTextBox.Text.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string relativePath = Convert.ToString(_assemblyList.SelectedItem);
            string assemblyPath = Path.GetFullPath(Path.Combine(rootPath, relativePath));

            try
            {
                AppendLog("開始自動開啟組立：" + relativePath);
                bool includeNested = _readScopeComboBox.SelectedIndex == 1;
                string scopeName = includeNested ? "完整階層" : "主組立第一層";
                IList<string> occurrences = SolidEdgeReader.OpenReadAndCloseAssembly(assemblyPath, includeNested);
                AppendLog("讀取完成（" + scopeName + "）：取得 " + occurrences.Count + " 個參照。");
                foreach (string item in occurrences)
                    AppendLog(item);

                AppendLog("自動開啟／讀取／關閉測試完成。原始檔未修改、未儲存。");
                MessageBox.Show(this,
                    "測試完成。\r\n組立：" + relativePath +
                    "\r\n讀取範圍：" + scopeName + "\r\n參照數：" + occurrences.Count +
                    "\r\n\r\n若此文件原本未開啟，工具已在讀取後以不儲存方式關閉。\r\n若文件原本已開啟，工具會保留開啟狀態。",
                    "自動開啟測試完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                string details = GetExceptionDetails(ex);
                AppendLog("自動開啟測試失敗：" + details);
                MessageBox.Show(this, details, "自動開啟測試失敗（診斷資訊）", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
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
