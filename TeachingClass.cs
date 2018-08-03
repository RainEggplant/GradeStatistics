using System.Collections.Generic;

namespace GradeStatistics
{
    class TeachingClass
    {
        private Dictionary<int, int> matchList = new Dictionary<int, int>();

        /// <summary>
        /// 获取或设置班级名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取班级的总权值。
        /// </summary>
        public double TotalWeight
        {
            get
            {
                double totalWeight = 0;
                foreach (int ruleIndex in matchList.Keys)
                {
                    totalWeight += MainWindow.RuleList[ruleIndex].Weight * matchList[ruleIndex];
                }
                return totalWeight;
            }
        }

        /// <summary>
        /// 获取该班级的详细统计信息。
        /// </summary>
        public string DetailedStatistics
        {
            get
            {
                string text = string.Empty;
                text += ("班级：" + Name + "\n");
                foreach (int ruleIndex in matchList.Keys)
                {
                    text += string.Format("\t{0}\t人数：{1}.\n",
                        MainWindow.RuleList[ruleIndex].ToString(),
                        matchList[ruleIndex].ToString());
                }
                text += "总权值：" + TotalWeight.ToString() + "\n";
                return text;
            }
        }

        /// <summary>
        /// 初始化 TeachingClass 类的新实例。
        /// </summary>
        /// <param name="name">班级名称</param>
        public TeachingClass(string name)
        {
            Name = name;
        }

        /// <summary>
        /// 设置某个规则下的匹配数量。
        /// </summary>
        /// <param name="ruleIndex">规则的索引</param>
        /// <param name="count">匹配数量</param>
        public void AddMatch(int ruleIndex, int count)
        {
            matchList.Add(ruleIndex, count);
        }

    }
}
