using System;
using System.Text;
using System.Windows.Forms;
using System.IO.Compression;
using System.IO;
using System.Diagnostics;
using MiniExcelLibs;
using MiniExcelLibs.Utils;
using System.Collections.Generic;

namespace hyyn_deploy_tool
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private string tempDir = "E:\temp";

        private void ExpButton_Click(object sender, EventArgs e)
        {
            if (CheckSource())
            {
                //开始导出数据
                UpdateMsg("开始导出数据...",10);
                string username = userNameTextBox.Text;
                if (username == "")
                {
                    MessageBox.Show("请输入用户名！");
                    progressBar1.Value = 0;
                    return;
                }
                string password = passwdTextBox.Text;
                if (password == "")
                {
                    MessageBox.Show("请输入密码！");
                    progressBar1.Value = 0;
                    return;
                }
                string dbName = dbNameTextBox.Text;
                if (dbName == "")
                {
                    MessageBox.Show("请输入数据库名！");
                    progressBar1.Value = 0;
                    return;
                }
                string sql = sqlTextBox.Text;
                if (sql == "")
                {
                    MessageBox.Show("请输入SQL！");
                    progressBar1.Value = 0;
                    return;
                }
                //获取当前系统登录用户名用于命名临时文件
                string csvFileName = Environment.UserName + DateTime.Now.ToString("yyyyMMddHH24mmss") + ".csv";
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
                string excelFileName = Environment.UserName + DateTime.Now.ToString("yyyyMMddHH24mmss") + ".xlsx";
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

        private Boolean CheckSource()
        {
            //检查临时目录是否存在
            if (!System.IO.Directory.Exists("E:/temp"))
            {
                UpdateMsg("临时目录不存在,开始生成。");
                //创建临时目录
                try
                {
                    System.IO.Directory.CreateDirectory("E:/temp");
                }
                catch (Exception ex)
                {
                    UpdateMsg("临时目录创建失败，请检查！");
                    UpdateMsg(ex.Message);
                    return false;
                }
            }
            //检查sqluldr264是否存在
            if (!System.IO.File.Exists("E:/temp/sqluldr264.exe"))
            {
                UpdateMsg("sqluldr264.exe不存在，开始生成！");
                // 解压sqluldr264.zip文件到E:/temp目录下
                try
                {
                    //加载source目录下的sqluldr264.zip文件到E:/temp目录下
                    string sourcePath = Path.Combine(Application.StartupPath, "source\\sqluldr264.zip");
                    if (!System.IO.File.Exists(sourcePath))
                    {
                        UpdateMsg("sqluldr264.zip资源文件不存在，请检查！");
                        return false;
                    }
                    System.IO.File.Copy(sourcePath, "E:/temp/sqluldr264.zip");
                    string zipFilePath = "E:/temp/sqluldr264.zip";
                    string extractPath = "E:/temp";
                    ZipFile.ExtractToDirectory(zipFilePath, extractPath);
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
                FileStream csv = File.Open(Path.Combine(tempDir, csvFileName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                FileStream xlsx = new FileStream(Path.Combine(tempDir, excelFileName), FileMode.CreateNew);
                IEnumerable<object> value = csv.Query(useHeaderRow: false, null, ExcelType.CSV, configuration: readConfig);
                xlsx.SaveAs(value, printHeader: false,configuration:writeConfig);
                //MiniExcel.ConvertCsvToXlsx(Path.Combine(tempDir, csvFileName), Path.Combine(tempDir, excelFileName));
                xlsx.Close();
                csv.Close();
                UpdateTextBox(logTextBox, "转换CSV数据为Excel文件完成！");
                finish(logTextBox, progressBar1, "导出Excel数据完成！文件路径：\r"
                    + Path.Combine(tempDir, excelFileName));
            };
            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            //process.WaitForExit();
            
            //process.Dispose();
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
        private void finish(RichTextBox textBox,ProgressBar progressBar, string msg)
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
            string time = DateTime.Now.ToString("yyyy年MM月dd日ttHH时mm分ss秒");
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

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //设置临时目录，默认为E:/temp
            folderBrowserDialog1.ShowDialog();
            tempDir = folderBrowserDialog1.SelectedPath;
            if (tempDir == "" || tempDir == null)
            {
                tempDir = "E:/temp";
            }
            
        }
    }
}
