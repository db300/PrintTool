using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordPrintTool
{
    public partial class MainForm : Form
    {
        #region constructor
        public MainForm()
        {
            InitializeComponent();

            InitUi();

            Load += MainForm_Load;
            FormClosing += MainForm_FormClosing;
        }
        #endregion

        #region property
        private TextBox _txtPath;
        private TextBox _txtLog;
        private Microsoft.Office.Interop.Word.Application _word;
        #endregion

        #region event handler
        private void MainForm_Load(object sender, EventArgs e)
        {
            _word = new Microsoft.Office.Interop.Word.Application { Visible = false };
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _word.Quit(SaveChanges: false);
        }

        private void BtnBrowser_Click(object sender, EventArgs e)
        {
            var dlg = new FolderBrowserDialog
            {
                Description = "选择要打印的目录",
                //RootFolder = Environment.SpecialFolder.MyComputer
                SelectedPath = Application.StartupPath
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            _txtPath.Text = dlg.SelectedPath;
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            var files = Directory.GetFiles(_txtPath.Text, "*.docx").ToList();
            if (!(files?.Count > 0))
            {
                _txtLog.AppendText($"{_txtPath} 目录下无docx文件\r\n");
                return;
            }

            var background = new BackgroundWorker { WorkerReportsProgress = true };
            background.DoWork += (ww, ee) =>
            {
                if (!(ee.Argument is Microsoft.Office.Interop.Word.Application word)) return;
                for (var i = 0; i < files.Count; i++)
                {
                    var file = files[i];
                    try
                    {
                        var doc = word.Documents.Open(file, ReadOnly: true, Visible: Missing.Value);
                        doc.PrintOut();
                        doc.Close(SaveChanges: false);
                        background.ReportProgress(i, $"{file} 打印...");
                    }
                    catch (Exception ex)
                    {
                        background.ReportProgress(i, $"{file} {ex.Message}");
                    }
                }
            };
            background.ProgressChanged += (ww, ee) =>
            {
                _txtLog.AppendText($"【{ee.ProgressPercentage} / {files.Count}】{ee.UserState}\r\n");
            };
            background.RunWorkerCompleted += (ww, ee) =>
            {
                _txtLog.AppendText("打印完成\r\n");
            };
            background.RunWorkerAsync(_word);
            _txtLog.AppendText("开始打印\r\n");
        }
        #endregion

        #region ui
        private void InitUi()
        {
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Word打印工具";

            var btnBrowser = new Button
            {
                AutoSize = true,
                Location = new Point(10, 10),
                Parent = this,
                Text = "选择目录"
            };
            btnBrowser.Click += BtnBrowser_Click;

            var btnPrint = new Button
            {
                AutoSize = true,
                Location = new Point(btnBrowser.Right + 10, btnBrowser.Top),
                Parent = this,
                Text = "打印"
            };
            btnPrint.Click += BtnPrint_Click;

            _txtPath = new TextBox
            {
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right,
                Location = new Point(btnBrowser.Left, btnBrowser.Bottom + 10),
                Parent = this,
                Width = ClientSize.Width - 20
            };

            _txtLog = new TextBox
            {
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom,
                Location = new Point(_txtPath.Left, _txtPath.Bottom + 10),
                Multiline = true,
                Parent = this,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Size = new Size(_txtPath.Width, ClientSize.Height - _txtPath.Bottom - 20),
                WordWrap = false
            };
        }
        #endregion
    }
}
