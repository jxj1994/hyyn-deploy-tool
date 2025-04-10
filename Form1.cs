using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using MiniExcelLibs;
using System.Collections.Generic;
using System.IO.Compression;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace hyyn_deploy_tool
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private string tempDir = "G:\\temp";

        private void ExpButton_Click(object sender, EventArgs e)
        {
            if (CheckSource())
            {
                expButton.Enabled = false;
                //开始导出数据
                UpdateMsg("开始导出数据...",10);
                string username = userNameTextBox.Text;
                if (username == "")
                {
                    MessageBox.Show("请输入用户名！");
                    progressBar1.Value = 0;
                    expButton.Enabled = true;
                    return;
                }
                string password = passwdTextBox.Text;
                if (password == "")
                {
                    MessageBox.Show("请输入密码！");
                    progressBar1.Value = 0;
                    expButton.Enabled = true;
                    return;
                }
                string dbName = dbNameTextBox.Text;
                if (dbName == "")
                {
                    MessageBox.Show("请输入数据库连接信息！");
                    progressBar1.Value = 0;
                    expButton.Enabled = true;
                    return;
                }
                else
                {
                    SaveConnection(dbName);
                }
                    string sql = sqlTextBox.Text;
                if (sql == "")
                {
                    MessageBox.Show("请输入SQL！");
                    progressBar1.Value = 0;
                    expButton.Enabled = true;
                    return;
                }
                //获取当前系统登录用户名用于命名临时文件
                string csvFileName = Environment.UserName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
                StringBuilder command = new StringBuilder();
                command.Append(Path.Combine(tempDir, "sqluldr264.exe"))
                .Append(" user=")
                .Append(username)
                .Append("/")
                .Append(password)
                .Append("@")
                .Append(dbName)
                .Append(" query=\"")
                .Append(sql)
                .Append("\" rows=1000 file=")
                .Append(Path.Combine(tempDir,csvFileName))
                .Append(" head=yes text=csv");
                //执行命令行命令导出数据
                UpdateMsg("开始执行导出命令...\r"+command.ToString(),20);
                string excelFileName = Environment.UserName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                try
                {
                    RunCommand(csvFileName, excelFileName,command.ToString());
                }
                catch (Exception ex)
                {
                    UpdateMsg("导出数据失败，请检查！");
                    UpdateMsg(ex.Message);
                    return;
                }
            }
            else
            {
                MessageBox.Show("资源文件初始化失败！");
            }
        }


        private bool CheckSource()
        {
            //检查临时目录是否存在
            if (!Directory.Exists(tempDir))
            {
                UpdateMsg("临时目录不存在,开始生成。");
                //创建临时目录
                try
                {
                    Directory.CreateDirectory(tempDir);
                }
                catch (Exception ex)
                {
                    UpdateMsg("临时目录创建失败，请检查！");
                    UpdateMsg(ex.Message);
                    return false;
                }
            }
            //检查sqluldr264是否存在
            if (!File.Exists(Path.Combine(tempDir,"sqluldr264.exe")))
            {
                UpdateMsg("sqluldr264.exe不存在，开始生成！");
                // 解压sqluldr264.zip文件到E:/temp目录下
                try
                {
                    //加载source目录下的sqluldr264.zip文件到E:/temp目录下
                    string sourcePath = Path.Combine(Application.StartupPath, "source\\sqluldr264.zip");
                    if (!File.Exists(sourcePath))
                    {
                        UpdateMsg("sqluldr264.zip资源文件不存在，请检查！");
                        return false;
                    }
                    File.Copy(sourcePath, Path.Combine(tempDir, "sqluldr264.zip"));
                    string zipFilePath = Path.Combine(tempDir, "sqluldr264.zip");
                    ZipFile.ExtractToDirectory(zipFilePath, tempDir);
                    File.Delete(zipFilePath);
                }
                catch (Exception ex)
                {
                    UpdateMsg("sqluldr264.exe解压失败，请检查！");
                    UpdateMsg(ex.Message);
                    return false;
                }
            }
            return true;
        }

        private void RunCommand(string csvFileName,string excelFileName, string command)
        {
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c {command}",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            process.StartInfo = startInfo;
            process.OutputDataReceived += (sender, e) => UpdateTextBox(logTextBox, e.Data);
            process.ErrorDataReceived += (sender, e) => UpdateTextBox(logTextBox, $"[Error] {e.Data}");
            // 订阅Exited事件
            process.EnableRaisingEvents = true;  // 必须启用事件
            process.Exited += (sender, e) =>
            {
                // 进程退出后的操作
                // 此处可调用方法或更新UI（需跨线程处理）
                UpdateTextBox(logTextBox, "导出CSV数据完成！");
                UpdateTextBox(logTextBox, "开始转换CSV数据为Excel文件...");
                var readConfig = new MiniExcelLibs.Csv.CsvConfiguration()
                {
                    StreamReaderFunc = (stream) => new StreamReader(stream, Encoding.GetEncoding("gb2312"))
                };
                var writeConfig = new MiniExcelLibs.Csv.CsvConfiguration()
                {
                    StreamWriterFunc = (stream) => new StreamWriter(stream, Encoding.GetEncoding("gb2312"))
                };
                if (File.Exists(Path.Combine(tempDir, csvFileName)))
                {
                    FileStream csv = File.Open(Path.Combine(tempDir, csvFileName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    FileStream xlsx = new FileStream(Path.Combine(tempDir, excelFileName), FileMode.CreateNew);
                    IEnumerable<object> value = csv.Query(useHeaderRow: false, null, ExcelType.CSV, configuration: readConfig);
                    xlsx.SaveAs(value, printHeader: false, configuration: writeConfig);
                    xlsx.Close();
                    csv.Close();
                    UpdateTextBox(logTextBox, "转换CSV数据为Excel文件完成！");
                    Finish(logTextBox, progressBar1, "导出Excel数据完成！文件路径：\r"
                        + Path.Combine(tempDir, excelFileName));
                }
                else
                {
                    MessageBox.Show("导出Excel失败，临时目录" + tempDir + "中无法找到csv文件");
                    Finish(logTextBox, progressBar1, "导出Excel失败，临时目录"+tempDir+"中无法找到csv文件");
                }
                File.Delete(Path.Combine(tempDir, csvFileName));
                if (expButton.InvokeRequired) {
                    expButton.Invoke(new Action(() =>
                    {
                        expButton.Enabled = true;
                    }));
                }
                else
                {
                    expButton.Enabled = true;
                }
            };
            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
        }

        private void UpdateTextBox(RichTextBox textBox, string data)
        {
            if (string.IsNullOrEmpty(data)) return;
            if (textBox.InvokeRequired)
            {
                textBox.Invoke(new Action(() =>
                {
                    textBox.AppendText(data + Environment.NewLine);
                    textBox.ScrollToCaret();
                }));
            }
            else
            {
                textBox.AppendText(data + Environment.NewLine);
                textBox.ScrollToCaret();
            }
        }

        /**
         * 结束的操作
         */
        private void Finish(RichTextBox textBox,ProgressBar progressBar, string msg)
        {
            if (string.IsNullOrEmpty(msg)) return;
            if (textBox.InvokeRequired)
            {
                textBox.Invoke(new Action(() =>
                {
                    textBox.AppendText(msg + Environment.NewLine);
                    textBox.ScrollToCaret();
                }));
            }
            else
            {
                textBox.AppendText(msg + Environment.NewLine);
                textBox.ScrollToCaret();
            }
            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() =>
                {
                    progressBar.Value = 100;
                }));
            }
            else
            {
                progressBar.Value = 100;
            }
        }


        private void UpdateMsg(string msg)
        {
            //获取当前时间进行日志输出
            string time = DateTime.Now.ToString("yyyy年MM月dd日HH时mm分ss秒");
            logTextBox.AppendText(time + " ");
            logTextBox.AppendText(msg + "\r");
            logTextBox.ScrollToCaret();
        }

        private void UpdateMsg(string msg,int process)
        {
            UpdateMsg(msg);
            progressBar1.Value = process;
        }


        /*
        * 退出
        */
        private void QuitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 退出程序
            Application.Exit();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //设置临时目录，默认为E:/temp
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                tempDir = dialog.FileName;
            }
            if (tempDir == "" || tempDir == null)
            {
                tempDir = "G:\\temp";
            }
            UpdateMsg("临时目录被设置为：" + tempDir);
            
        }


        private void SaveConnection(string dbName)
        {
            File.AppendAllLines(Path.Combine(Application.StartupPath, "connections.txt"), new string[] { dbName });
        }

        /**
         * 连接历史选择
         */
        private void ToolStripComboBox1_Click(object sender, EventArgs e)
        {
            //连接历史保存为json文件，加载json文件作为选择项
            if (File.Exists(Path.Combine(Application.StartupPath,"connections.txt")))
            {
                foreach (var item in File.ReadLines(Path.Combine(Application.StartupPath, "connections.txt")))
                {
                    if (toolStripComboBox1.Items.Count == 0)
                    {
                        toolStripComboBox1.Items.Add(item);
                    }
                    else
                    {
                        foreach (var item1 in toolStripComboBox1.Items)
                        {
                            if (!item1.ToString().Equals(item.ToString()))
                            {
                                toolStripComboBox1.Items.Add(item);
                            }
                        }
                    }
                }
                
            }
        }

        private void ToolStripComboBox1_SelectedItemChanged(object sender, EventArgs e)
        {
            dbNameTextBox.Text = toolStripComboBox1.SelectedItem.ToString();
            mainToolStripMenuItem.HideDropDown();
        }


    }
}
