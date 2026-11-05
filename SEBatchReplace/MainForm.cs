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
            var root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.Padding = new Padding(12);
            root.ColumnCount = 1;
            root.RowCount = 5;
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 34));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 18));
            Controls.Add(root);

            var folderPanel = new Panel { Dock = DockStyle.Fill };
            var folderLabel = new Label { Text = "組立資料夾（只掃描第一層 .asm）", AutoSize = true, Location = new Point(0, 0) };
            _folderTextBox = new TextBox { Location = new Point(0, 25), Width = 820, Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right };
            var browseButton = new Button { Text = "選擇資料夾", Location = new Point(830, 23), Width = 120, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            browseButton.Click += BrowseFolder;
            folderPanel.Controls.Add(folderLabel);
            folderPanel.Controls.Add(_folderTextBox);
            folderPanel.Controls.Add(browseButton);
            root.Controls.Add(folderPanel, 0, 0);

            var assemblyGroup = new GroupBox { Text = "找到的組立檔", Dock = DockStyle.Fill };
            _assemblyList = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true };
            assemblyGroup.Controls.Add(_assemblyList);
            root.Controls.Add(assemblyGroup, 0, 1);

            var mappingGroup = new GroupBox { Text = "手動指定零件對應（舊 .par → 新 .par）", Dock = DockStyle.Fill };
            var mappingLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(6) };
            mappingLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            mappingLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            _mappingList = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true };
            var mappingButtons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            var addButton = new Button { Text = "新增配對", Width = 110, Height = 32 };
            addButton.Click += AddMapping;
            var removeButton = new Button { Text = "移除配對", Width = 110, Height = 32 };
            removeButton.Click += RemoveMapping;
            mappingButtons.Controls.Add(addButton);
            mappingButtons.Controls.Add(removeButton);
            mappingLayout.Controls.Add(_mappingList, 0, 0);
            mappingLayout.Controls.Add(mappingButtons, 1, 0);
            mappingGroup.Controls.Add(mappingLayout);
            root.Controls.Add(mappingGroup, 0, 2);

            var policyGroup = new GroupBox { Text = "遇到關聯失效時的處理方式（執行前選擇）", Dock = DockStyle.Fill };
            _skipRadio = new RadioButton { Text = "回復該次嘗試、跳過並繼續後續替換", AutoSize = true, Location = new Point(12, 22), Checked = true };
            _stopRadio = new RadioButton { Text = "回復該次嘗試、立即停止批次作業", AutoSize = true, Location = new Point(12, 46) };
            policyGroup.Controls.Add(_skipRadio);
            policyGroup.Controls.Add(_stopRadio);
            root.Controls.Add(policyGroup, 0, 3);

            var logGroup = new GroupBox { Text = "作業紀錄", Dock = DockStyle.Fill };
            _logTextBox = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
            logGroup.Controls.Add(_logTextBox);
            root.Controls.Add(logGroup, 0, 4);

            AppendLog("目前為介面與掃描骨架，尚未連接 Solid Edge，也不會修改任何組立或零件檔。");
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
                    foreach (string file in files)
                        _assemblyList.Items.Add(file);
                    AppendLog("掃描完成：找到 " + files.Count + " 個第一層 .asm 檔案。");
                }
                catch (Exception ex)
                {
                    AppendLog("掃描失敗：" + ex.Message);
                    MessageBox.Show(this, ex.Message, "掃描失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void AddMapping(object sender, EventArgs e)
        {
            string oldPath = SelectPartFile("選擇舊零件（.par）");
            if (oldPath == null) return;
            string newPath = SelectPartFile("選擇新零件（.par）");
            if (newPath == null) return;

            if (String.Equals(oldPath, newPath, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this, "新舊零件不可為同一個檔案。", "配對無效", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (ReplacementMapping mapping in _mappings)
            {
                if (String.Equals(mapping.OldPartPath, oldPath, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(this, "這個舊零件已經建立配對。", "重複配對", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            _mappings.Add(new ReplacementMapping(oldPath, newPath));
            _mappingList.Items.Add(_mappings[_mappings.Count - 1]);
            AppendLog("新增配對：" + Path.GetFileName(oldPath) + " → " + Path.GetFileName(newPath));
        }

        private string SelectPartFile(string title)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = title;
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
