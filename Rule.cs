using System;

namespace GradeStatistics
{
    class Rule
    {
        /// <summary>
        /// 获取或设置分数下限。
        /// </summary>
        public int LBound { get; set; }

        /// <summary>
        /// 获取或设置分数上限。
        /// </summary>
        public int UBound { get; set; }

        /// <summary>
        /// 获取或设置权值。
        /// </summary>
        public double Weight { get; set; }


        /// <summary>
        /// 初始化 Rule 类的新实例。
        /// </summary>
        /// <param name="lBound">分数下限</param>
        /// <param name="uBound">分数上限</param>
        /// <param name="weight">权值</param>
        public Rule(int lBound = 0, int uBound = 0, double weight = 0)
        {
            if (lBound < 0 || uBound < 0)
                throw new ArgumentOutOfRangeException("分数上下限不能为负值！");
            if (lBound > uBound)
                throw new ArgumentOutOfRangeException("分数下限，分数上限", "分数下限不能大于分数上限！");            
            LBound = lBound;
            UBound = uBound;
            Weight = weight;
        }


        /// <summary>
        /// 获取该规则的文字表述形式。
        /// </summary>
        public override string ToString()
        {
            return string.Format("分数段 {0, -15}权值 {1, -6}",
                string.Format("[{0}, {1}],", LBound.ToString(), UBound.ToString())
                , Weight.ToString() + ".");
        }

        /// <summary>
        /// 获取该规则对应的 SQL 语句。
        /// </summary>
        /// <param name="sheet">工作表名称</param>
        /// <param name="classColumn">班级列名称</param>
        /// <param name="queryClass">待统计班级的名称</param>
        /// <param name="scoreColumn">总分列名称</param>
        /// <returns>该规则对应的 SQL 语句</returns>
        public string QueryCommand(
            string sheet, string classColumn, string queryClass, string scoreColumn)
        {
            return string.Format(
                @"SELECT COUNT(*) FROM [{0}$] WHERE {1}='{2}' AND (CInt({3}) BETWEEN {4} AND {5});",
                sheet, classColumn, queryClass, scoreColumn, LBound, UBound);
        }
    }
}
