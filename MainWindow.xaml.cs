using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace GradeStatistics
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        // 定义连接至 Excel 文件的所需对象。
        private OleDbConnection dbConnection;
        private OleDbCommand dbCommand;
        // 工作表和列列表。
        private List<string> sheetList;
        private List<string> columnList;
        // 是否已加载的标识。
        private bool isLoaded = false;
        // 规则列表。
        internal static List<Rule> RuleList = new List<Rule>();
        private ObservableCollection<string> strRuleList = new ObservableCollection<string>();
        internal static List<TeachingClass> classList;


        public MainWindow()
        {
            InitializeComponent();
            lstRule.ItemsSource = strRuleList;
        }

        private void mnuOpen_Click(object sender, RoutedEventArgs e)
        {
            // 用户选择要打开的 Excel 文件。
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                CheckFileExists = true,
                Filter = "Excel 文件|*.xlsx;*.xls"
            };
            if (openFileDialog.ShowDialog() == false)
                return;

            string connectionString;
            string filename = openFileDialog.FileName;
            // 根据 Excel 文件版本选择对应的 ConnectionString.
            switch (System.IO.Path.GetExtension(filename).ToLower())
            {
                case ".xls":
                    connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source="
                        + filename + "; Extended Properties=\"Excel 8.0; HDR=NO; IMEX=1\";";
                    break;
                case ".xlsx":
                    connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source="
                        + filename + "; Extended Properties=\"Excel 12.0; HDR=NO; IMEX=1\";";
                    break;
                default:
                    MessageBox.Show("错误的文件格式！请检查您的文件。", "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
            }
            // 连接到 Excel 文件并获取其包含的工作表。
            if (dbConnection != null)
            {
                dbConnection.Dispose();
                isLoaded = false;
                btnExecute.IsEnabled = false;
            }
            try
            {
                dbConnection = new OleDbConnection(connectionString);
                dbConnection.Open();
                DataTable schemaTable = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                sheetList = new List<string>();
                foreach (DataRow row in schemaTable.Rows)
                {
                    if (row["TABLE_NAME"].ToString().EndsWith("$"))
                        sheetList.Add(row["TABLE_NAME"].ToString().TrimEnd('$'));
                }
                // sheetList.Sort((x, y) => ExtractNumber(x).CompareTo(ExtractNumber(y)));
                dbCommand = dbConnection.CreateCommand();
                // 更新 UI.
                btnExecute.IsEnabled = true;
                lblFilename.Content = openFileDialog.SafeFileName;
                cmbSheets.ItemsSource = sheetList;
                isLoaded = true;
                cmbSheets.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException)
                {
                    if (MessageBox.Show(ex.Message + "\n您可能需要下载并安装 " +
                        "\"2007 Office system 驱动程序：数据连接组件\"。\n" +
                        "是否前往下载？", "错误",
                        MessageBoxButton.YesNo, MessageBoxImage.Error) == MessageBoxResult.Yes)
                    {
                        System.Diagnostics.Process.Start(
                            "https://www.microsoft.com/zh-cn/download/details.aspx?id=23734");
                    }
                }
                else
                {
                    MessageBox.Show(ex.Message, "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
                if (dbConnection != null)
                    dbConnection.Dispose();
                isLoaded = false;
                return;
            }
        }

        private void mnuExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void cmbSheets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!isLoaded)
                return;

            // 显示选择的工作表的预览。
            dbCommand.CommandText = string.Format(
                @"SELECT TOP 10 * FROM [{0}$];", cmbSheets.SelectedItem);
            OleDbDataAdapter dbDataAdapter = new OleDbDataAdapter(dbCommand);
            DataSet excelDataSet = new DataSet();
            dbDataAdapter.Fill(excelDataSet);
            dgView.ItemsSource = excelDataSet.Tables[0].AsDataView();
            // 获取所有列的名称。
            DataTable schemaTable = dbConnection.GetOleDbSchemaTable(
                OleDbSchemaGuid.Columns, new Object[] { null, null, cmbSheets.SelectedItem + "$" });
            columnList = new List<string>();
            foreach (DataRow row in schemaTable.Rows)
            {
                columnList.Add(row["COLUMN_NAME"].ToString());
            }
            columnList.Sort((x, y) => ExtractNumber(x).CompareTo(ExtractNumber(y)));
            // 更新 UI.
            cmbClass.ItemsSource = columnList;
            cmbScore.ItemsSource = columnList;
            cmbClass.SelectedIndex = 0;
            cmbScore.SelectedIndex = 0;
        }

        private void btnAddRule_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                RuleList.Add(new Rule(
                    int.Parse(txtLBound.Text), int.Parse(txtUBound.Text),
                    double.Parse(txtWeight.Text)));
                strRuleList.Add(RuleList.Last().ToString());
                lstRule.SelectedIndex = lstRule.Items.Count - 1;
                txtLBound.Text = string.Empty;
                txtUBound.Text = string.Empty;
                txtWeight.Text = string.Empty;
                txtLBound.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnDeleteRule_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                RuleList.RemoveAt(lstRule.SelectedIndex);
                strRuleList.RemoveAt(lstRule.SelectedIndex);
                lstRule.SelectedIndex = lstRule.Items.Count - 1;
            }
            catch (ArgumentOutOfRangeException)
            { }
        }

        private void btnExecute_Click(object sender, RoutedEventArgs e)
        {
            classList = new List<TeachingClass>();
            string sheetName = cmbSheets.SelectedItem.ToString();
            string classColumn;
            string scoreColumn;
            try
            {
                classColumn = cmbClass.SelectedItem.ToString();
                scoreColumn = cmbScore.SelectedItem.ToString();
                // 若选择了同一列，则抛出异常。
                if (cmbClass.SelectedIndex == cmbScore.SelectedIndex)
                    throw new Exception(
                        "班级和成绩不能是同一列，请检查您的选择。\n" +
                        new string(' ', 50) + "——“心不静”");
                // 若规则为空，则抛出异常。
                if (RuleList.Count == 0)
                    throw new Exception("规则不能为空！请至少指定一条规则。");
                // 获取表中所有班级。
                dbCommand.CommandText = string.Format(
                    @"SELECT DISTINCT {0} FROM [{1}$] WHERE {0} LIKE '%[0-9]';",
                    classColumn, cmbSheets.SelectedItem);
                using (OleDbDataReader dataReader = dbCommand.ExecuteReader())
                {
                    while (dataReader.Read())
                    {
                        classList.Add(new TeachingClass(dataReader[0].ToString()));
                    }
                }
                // 若未获取到班级，则抛出异常。
                if (classList.Count == 0)
                    throw new Exception("没有找到班级！请检查您的文件。");
                classList.Sort((x, y) => ExtractNumber(x.Name).CompareTo(ExtractNumber(y.Name)));
                // 根据规则进行统计。
                foreach (TeachingClass tClass in classList)
                {
                    for (int i = 0; i < RuleList.Count; ++i)
                    {
                        dbCommand.CommandText =
                            RuleList[i].QueryCommand(sheetName, classColumn, tClass.Name, scoreColumn);
                        using (OleDbDataReader dataReader = dbCommand.ExecuteReader())
                        {
                            dataReader.Read();
                            // 记录此班级匹配当前规则的人数。
                            tClass.AddMatch(i, int.Parse(dataReader[0].ToString()));
                        }
                    }
                }
                ResultWindow resultWindow = new ResultWindow();
                resultWindow.ShowDialog();
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("请选择班级和总分列的列名称！", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void mnuClearRule_Click(object sender, RoutedEventArgs e)
        {
            RuleList.Clear();
            strRuleList.Clear();
        }

        private void mnuImportRule_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                CheckFileExists = true,
                Filter = "规则文件|*.rule"
            };
            if (openFileDialog.ShowDialog() == false)
                return;

            int count = RuleList.Count;
            using (StreamReader sr = new StreamReader(openFileDialog.FileName))
            {
                while (!sr.EndOfStream)
                {
                    string[] rule = sr.ReadLine().Split(' ');
                    RuleList.Add(new Rule(int.Parse(rule[0]), int.Parse(rule[1]),
                        double.Parse(rule[2])));
                    strRuleList.Add(RuleList.Last().ToString());
                }
            }
            MessageBox.Show("成功导入 " + (RuleList.Count - count).ToString() + " 条规则！",
                "导入规则", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void mnuExportRule_Click(object sender, RoutedEventArgs e)
        {
            if (RuleList.Count == 0)
            {
                MessageBox.Show(
                    "规则为空时不能导出规则。\n" +
                    new string(' ', 30) + "——“低级错误”", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                AddExtension = true,
                Filter = "规则文件|*.rule"
            };
            if (saveFileDialog.ShowDialog() == false)
                return;

            using (StreamWriter sw = new StreamWriter(saveFileDialog.FileName))
            {
                foreach (Rule rule in RuleList)
                {
                    sw.WriteLine(string.Format("{0} {1} {2}",
                        rule.LBound.ToString(), rule.UBound.ToString(), rule.Weight.ToString()));
                }
            }
            MessageBox.Show("成功导出 " + RuleList.Count.ToString() + " 条规则！", "导出规则",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void txtWeight_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                btnAddRule_Click(null, null);
        }

        private void mnuHelp_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "使用说明：\n" +
                "    使用时，您需要先通过『文件』菜单打开要进行分析的 Excel 文档，本程序将显示文档的预览。" +
                "然后，您需要根据预览表格指定『工作表』、『班级列名称』和『总分列名称』。\n" +
                "    在右方，您可以添加统计规则，同时，您可以通过『规则』菜单导出或导入规则。小提示：输入规则" +
                "时，可通过『Tab』键切换文本框，『Enter』键添加规则，省去鼠标操作，减少『手部运动』。\n" +
                "    配置完毕后，请点击『执行统计』按钮。程序将弹出统计结果窗口，之后您可以保存更详细的" +
                "统计数据到文件中。",
                "帮助", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void mnuAbout_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(
                "成绩统计助手 V 1.0\n" +
                "    一个帮助你逃脱『手部运动』的小工具，-_-||\n\n" +
                "作者：RainEggplant ( https://www.raineggplant.com/ ) \n" +
                "是否访问作者的个人网站？",
                "关于", MessageBoxButton.YesNo, MessageBoxImage.Information)
                == MessageBoxResult.Yes)
            {
                System.Diagnostics.Process.Start(
                    "https://www.raineggplant.com/");
            }
        }

        /// <summary>
        /// 从文本中提取非负整数。
        /// </summary>
        /// <param name="text">文本</param>
        /// <returns>提取到的整数; 如果提取失败则返回 -1.</returns>
        private static int ExtractNumber(string text)
        {
            Match match = Regex.Match(text, @"\d+");
            if (match == null)
            {
                return -1;
            }
            if (!int.TryParse(match.Value, out int value))
            {
                return -1;
            }
            return value;
        }

    }
}
