using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using MiniExcelLibs;
using System.Collections.Generic;
using System.IO.Compression;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Text.Json.Nodes;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Oracle.ManagedDataAccess.Client;
using System.Linq;

namespace hyyn_deploy_tool
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private string tempDir = "G:\\temp";

        private string errorMsg = "";

        private Dictionary<string, ConnectInfo> connectInfoDict = new Dictionary<string, ConnectInfo>();

        private void ExpButton_Click(object sender, EventArgs e)
        {
            // TODO 
            /* TODO
             * 1.优化oracle导出工具，以60万行这个表格进行拆分
             * 2.当导出数据量太大时，数据会缺失，数据量太大时，进行拆分处置
             * 3.已连接成功的配置进行保存，下次使用选择对应的服务直接带出 */
            if (CheckSource())
            {
                expButton.Enabled = false;
                //开始导出数据
                UpdateMsg("开始导出数据...", 10);
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
                    SaveConnection(dbName, username);
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
                command.Append(Path.Combine(Application.StartupPath, "source", "sqluldr264.exe"))
                .Append(" user=")
                .Append(username)
                .Append("/")
                .Append(password)
                .Append("@")
                .Append(dbName)
                .Append(" sql=\"")
                .Append(sql)
                .Append("\" rows=1000 file=")
                .Append(Path.Combine(tempDir, csvFileName))
                .Append(" head=yes text=csv");
                //执行命令行命令导出数据
                UpdateMsg("开始执行导出命令...\r" + command.ToString(), 20);
                string excelFileName = Environment.UserName + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                try
                {
                    RunCommand(csvFileName, excelFileName, command.ToString());
                }
                catch (Exception ex)
                {
                    UpdateMsg("导出数据失败，请检查！");
                    UpdateMsg(ex.Message);
                    return;
                }
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
                    UpdateMsg("临时目录不可用，请选择其他临时目录！");
                    UpdateMsg(ex.Message);
                    return false;
                }
            }
            //检查sqluldr264是否存在
            if (!File.Exists(Path.Combine(Application.StartupPath, "source", "sqluldr264.exe")))
            {
                UpdateMsg("依赖文件source/sqluldr264.exe不存在，无法导出！");
                expButton.Enabled = false;
                return false;
            }
            return true;
        }

        private void RunCommand(string csvFileName, string excelFileName, string command)
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
                    Finish(logTextBox, progressBar1, "导出Excel失败，临时目录" + tempDir + "中无法找到csv文件");
                }
                File.Delete(Path.Combine(tempDir, csvFileName));
                if (expButton.InvokeRequired)
                {
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
        private void Finish(RichTextBox textBox, ProgressBar progressBar, string msg)
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

        private void UpdateMsg(string msg, int process)
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
            //连接历史保存为json文件，加载json文件作为选择项
            if (File.Exists(Path.Combine(Application.StartupPath, "connections.json")))
            {
                // 从文件中读取json数组
                string jsonText = File.ReadAllText(Path.Combine(Application.StartupPath, "connections.json"));
                List<ConnectInfo> connectInfos = JsonSerializer.Deserialize<List<ConnectInfo>>(jsonText);
                for (int i = 0; i < connectInfos.Count; i++)
                {
                    ConnectInfo item = connectInfos[i];
                    connectInfoDict.Add((i + 1) + "：" + item.user + "@" + item.dbName, item);
                    if (toolStripComboBox1.Items.Count == 0)
                    {
                        toolStripComboBox1.Items.Add((i + 1) + "：" + item.user + "@" + item.dbName);
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
            if (CheckSource())
            {
                UpdateMsg("资源文件初始化成功！");
            }
            else
            {
                UpdateMsg("资源文件初始化失败！");
            }
        }

        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //设置临时目录，默认为E:/temp
            CommonOpenFileDialog dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true
            };
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

        /**
         * 保存连接信息
         */
        private void SaveConnection(string dbName, string user)
        {
            ConnectInfo connectInfo = new ConnectInfo(dbName, user);
            List<ConnectInfo> connectInfoList = new List<ConnectInfo>();
            // 从文件中读取json数组
            if (File.Exists(Path.Combine(Application.StartupPath, "connections.json")))
            {
                try
                {
                    string jsonText = File.ReadAllText(Path.Combine(Application.StartupPath, "connections.json"));
                    // [{}]
                    JsonArray jsonArray = JsonSerializer.Deserialize<JsonArray>(jsonText);
                    if (jsonArray.Count > 0)
                    {

                        connectInfoList = JsonSerializer.Deserialize<List<ConnectInfo>>(jsonText);
                    }

                }
                catch (Exception ex)
                {
                    UpdateMsg("读取历史连接失败，请检查文件格式是否正确" + ex.Message);
                }


            }
            // 判断是否已经存在
            foreach (ConnectInfo item in connectInfoList)
            {
                if (item.dbName.Equals(dbName) && item.user.Equals(user))
                {
                    continue;
                }
                else
                {
                    connectInfoList.Add(connectInfo);
                }

            }
            string testJson = JsonSerializer.Serialize(connectInfoList);
            File.WriteAllText(Path.Combine(Application.StartupPath, "connections.json")
                , JsonSerializer.Serialize(connectInfoList));
        }

        private void ToolStripComboBox1_SelectedItemChanged(object sender, EventArgs e)
        {
            dbNameTextBox.Text = toolStripComboBox1.SelectedItem.ToString();
            if (connectInfoDict.ContainsKey(toolStripComboBox1.SelectedItem.ToString()))
            {
                ConnectInfo connectInfo = connectInfoDict[toolStripComboBox1.SelectedItem.ToString()];
                dbNameTextBox.Text = connectInfo.dbName;
                userNameTextBox.Text = connectInfo.user;
            }
            mainToolStripMenuItem.HideDropDown();
        }

        private void TestBtn_Click(object sender, EventArgs e)
        {
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
            bool isSuccess = ConnectTest();
            if (isSuccess)
            {
                MessageBox.Show("连接成功！");
                SaveConnection(dbName, username);
            }
            else
            {
                MessageBox.Show("连接失败！\r" + errorMsg);
            }

        }

        /**
         * 数据库连接测试
         */
        private bool ConnectTest()
        {
            bool isSuccess = false;
            userNameTextBox.Text = userNameTextBox.Text.Trim();
            passwdTextBox.Text = passwdTextBox.Text.Trim();
            dbNameTextBox.Text = dbNameTextBox.Text.Trim();
            string username = userNameTextBox.Text;
            string password = passwdTextBox.Text;
            string dbName = dbNameTextBox.Text;
            //校验输入的数据库信息是否正确
            //使用正则表达式校验数据库信息是否正确
            if (!Regex.IsMatch(dbName, "[\\w.-]+:\\d+[/|:]\\w+$"))
            {
                this.errorMsg = "SB，数据库连接写错了！";
                expButton.Enabled = false;
                return false;
            }
            //打开一个数据库连接测试连接是否正常
            StringBuilder connectString = new StringBuilder();
            //"User Id=username;Password=password;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=hostname)(PORT=port))
            //(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=servicename)))";
            connectString.Append("User Id=")
                .Append(username)
                .Append(";Password=")
                .Append(password)
                .Append(";Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=")
                .Append(dbName.Split(':')[0])
                .Append(")(PORT=")
                .Append(dbName.Split(':')[1].Split('/')[0]);
            //如果数据库信息中包含/，则使用SERVICE_NAME连接，否则使用SID连接
            if (Regex.IsMatch(dbName, "[\\w.-]+:\\d+/\\w+$"))
            {
                connectString.Append("))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=")
                    .Append(dbName.Split('/')[1]);
            }
            else
            {
                connectString.Append("))(CONNECT_DATA=(SERVER=DEDICATED)(SID=")
                    .Append(dbName.Split(':')[2]);
            }
            connectString.Append(")))");
            using (OracleConnection conn = new OracleConnection(connectString.ToString()))
            {
                try
                {
                    conn.Open(); // 打开连接 [[6]]
                    Console.WriteLine("Connected successfully!");
                    // 执行查询
                    OracleCommand cmd = new OracleCommand("SELECT * FROM DUAL", conn);
                    //OracleDataReader reader = cmd.ExecuteReader();
                    //while (reader.Read())
                    //{
                    //    Console.WriteLine(reader["name"].ToString());
                    //}
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    isSuccess = true;
                }
                catch (Exception ex)
                {
                    UpdateMsg("Error: " + ex.Message);
                    this.errorMsg = ex.Message;
                }
                finally
                {
                    conn.Close();
                }
            }
            return isSuccess;
        }
    }
}
