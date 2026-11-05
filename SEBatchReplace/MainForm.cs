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
        private TextBox _logTextBox;
        private Button _revisionReadButton;

        public MainForm()
        {
            Text = "SE Batch Replace ST6 - Revision Manager 測試";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(820, 560);
            Size = new Size(1000, 680);
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
                RowCount = 4
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 42));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 58));
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

            var assemblyGroup = new GroupBox { Text = "找到的組立檔（包含子資料夾）", Dock = DockStyle.Fill };
            _assemblyList = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true };
            assemblyGroup.Controls.Add(_assemblyList);
            root.Controls.Add(assemblyGroup, 0, 1);

            var testPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 0)
            };
            _revisionReadButton = new Button
            {
                Text = "Revision Manager 讀取測試（全部組立）",
                Width = 260,
                Height = 32
            };
            _revisionReadButton.Click += RevisionManagerReadAllAssemblies;
            var hint = new Label
            {
                Text = "逐一讀取清單中的所有 .asm；唯讀，不修改、不儲存檔案。",
                AutoSize = true,
                Padding = new Padding(8, 8, 0, 0)
            };
            testPanel.Controls.Add(_revisionReadButton);
            testPanel.Controls.Add(hint);
            root.Controls.Add(testPanel, 0, 2);

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
            root.Controls.Add(logGroup, 0, 3);

            AppendLog("請選擇組立根資料夾，再執行 Revision Manager 讀取測試。");
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

                    AppendLog("資料夾掃描完成：包含所有子資料夾，共找到 " + files.Count + " 個 .asm 檔案。");
                    if (files.Count == 0)
                        AppendLog("指定資料夾內沒有找到 .asm 檔案。");
                }
                catch (Exception ex)
                {
                    _assemblyList.Items.Clear();
                    AppendLog("資料夾掃描失敗：" + ex.Message);
                    MessageBox.Show(this, ex.Message, "掃描失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
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
                current = current.InnerException;
                level++;
            }
            return details.ToString();
        }

        private void AppendLog(string message)
        {
            _logTextBox.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + Environment.NewLine);
        }
    }
}
