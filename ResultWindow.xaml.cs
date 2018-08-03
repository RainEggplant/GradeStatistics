using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace GradeStatistics
{
    /// <summary>
    /// ResultWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ResultWindow : Window
    {
        private List<string> resultList = new List<string>();

        public ResultWindow()
        {
            InitializeComponent();
            foreach (TeachingClass tClass in MainWindow.classList)
            {
                resultList.Add(string.Format("班级：{0, -10}总权值：{1}",
                    tClass.Name, tClass.TotalWeight));
            }
            lstResult.ItemsSource = resultList;
        }

        private void btnExportResult_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                AddExtension = true,
                Filter = "统计结果|*.txt",
                FileName = "result.txt"
            };
            if (saveFileDialog.ShowDialog() == false)
                return;

            using (StreamWriter sw = new StreamWriter(saveFileDialog.FileName))
            {
                foreach (TeachingClass tClass in MainWindow.classList)
                {
                    sw.WriteLine(tClass.DetailedStatistics);
                }
            }
            MessageBox.Show("成功导出了详细统计数据！", "导出",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

    }
}
