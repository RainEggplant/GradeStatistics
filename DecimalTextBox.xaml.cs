using System;
using System.Windows.Controls;

namespace GradeStatistics
{
    /// <summary>
    /// DecimalTextBox.xaml 的交互逻辑
    /// </summary>
    public partial class DecimalTextBox : TextBox
    {
        public DecimalTextBox()
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
                if (!Double.TryParse(textBox.Text.Replace(' ', 'x'), out double num))
                {
                    textBox.Text = textBox.Text.Remove(offset, change[0].AddedLength);
                    textBox.Select(offset, 0);
                }
            }
        }

    }
}
