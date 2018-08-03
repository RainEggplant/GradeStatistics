using System.Windows.Controls;
using System.Windows.Input;

namespace GradeStatistics
{
    /// <summary>
    /// NumericTextBox.xaml 的交互逻辑
    /// </summary>
    public partial class NumericTextBox : TextBox
    {
        public NumericTextBox()
        {
            InitializeComponent();
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // 屏蔽非法字符粘贴与输入。
            TextBox textBox = sender as TextBox;
            TextChange[] change = new TextChange[e.Changes.Count];
            e.Changes.CopyTo(change, 0);
            int offset = change[0].Offset;
            if (change[0].AddedLength > 0)
            {
                if (!int.TryParse(textBox.Text.Replace(' ', 'x'), out int num))
                {
                    textBox.Text = textBox.Text.Remove(offset, change[0].AddedLength);
                    textBox.Select(offset, 0);
                }
            }
        }

    }
}
