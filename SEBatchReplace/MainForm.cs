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
                Text = "組立資料夾（只掃描第一層 .asm）",
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
                Margin = new Padding(0, 0, 0, 0)
            };
            browseButton.Click += BrowseFolder;

            folderPanel.Controls.Add(folderLabel, 0, 0);
            folderPanel.SetColumnSpan(folderLabel, 2);
            folderPanel.Controls.Add(_folderTextBox, 0, 1);
            folderPanel.Controls.Add(browseButton, 1, 1);
            root.Controls.Add(folderPanel, 0, 0);

            var assemblyGroup = new GroupBox { Text = "找到的組立檔", Dock = DockStyle.Fill };
            _assemblyList = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true };
            assemblyGroup.Controls.Add(_assemblyList);
            root.Controls.Add(assemblyGroup, 0, 1);

            var mappingGroup = new GroupBox { Text = "手動指定零件對應（舊零件名稱 → 新 .par 檔案）", Dock = DockStyle.Fill };
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
