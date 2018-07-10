using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Model
{
    /// <summary>
    /// 成绩单
    /// </summary>
    public class TranscriptsEntity
    {
        /// <summary>
        /// 语文成绩
        /// </summary>
        public int ChineseScores { get; set; }

        /// <summary>
        /// 数学成绩
        /// </summary>
        public int MathScores { get; set; }
    }
}
