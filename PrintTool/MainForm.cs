using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PrintTool
{
    public partial class MainForm : Form
    {
        #region constructor

        public MainForm()
        {
            InitializeComponent();

            InitUi();
        }

        #endregion

        #region property

        private TreeView _tvFile;

        #endregion

        #region event handler

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            var openDlg = new OpenFileDialog { Multiselect = true };
            if (openDlg.ShowDialog() != DialogResult.OK) return;
            _tvFile.Nodes.Clear();
            _tvFile.Nodes.AddRange(openDlg.FileNames.Select(a => new TreeNode(a)).ToArray());
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            var background = new BackgroundWorker { WorkerReportsProgress = true, WorkerSupportsCancellation = true };
            background.DoWork += (ww, ee) =>
            {
                foreach (TreeNode node in _tvFile.Nodes)
                {
                    if (background.CancellationPending) break;
                    Print(node.Text);
                    background.ReportProgress(0, node);
                }
            };
            background.ProgressChanged += (ww, ee) =>
            {
                if (!(ee.UserState is TreeNode node)) return;
                node.Text += "【已发送...】";
            };
            background.RunWorkerCompleted += (ww, ee) => { };
            background.RunWorkerAsync();
        }

        #endregion

        #region method

        private void Print(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return;
            var info = new ProcessStartInfo(fileName)
            {
                CreateNoWindow = true,
                Verb = "Print",
                WindowStyle = ProcessWindowStyle.Hidden
            };
            Process.Start(info);
        }

        #endregion

        #region ui

        private void InitUi()
        {
            Text = "打印工具";

            var btnOpen = new Button
            {
                AutoSize = true,
                Location = new Point(10, 10),
                Parent = this,
                Text = "打开"
            };
            btnOpen.Click += BtnOpen_Click;
            var btnPrint = new Button
            {
                AutoSize = true,
                Location = new Point(btnOpen.Right + 10, 10),
                Parent = this,
                Text = "打印"
            };
            btnPrint.Click += BtnPrint_Click;
            _tvFile = new TreeView
            {
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom,
                ItemHeight = 25,
                Location = new Point(btnOpen.Left, btnOpen.Bottom + 10),
                Parent = this,
                Size = new Size(ClientSize.Width - 20, ClientSize.Height - 20 - btnOpen.Bottom)
            };
        }

        #endregion
    }
}
