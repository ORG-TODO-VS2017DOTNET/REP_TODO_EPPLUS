using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZhangJR.EPPlus
{
    /// <summary>
    /// ExcelRange 扩展方法
    /// </summary>
    public static class ExcelRangeExtend
    {
        /// <summary>
        /// 为 ExcelRange 添加字符串序列形式的数据验证
        /// </summary>
        public static void AddListDataValidation(this ExcelRange target, params string[] datas)
        {
            if (target == null)
            {
                throw new NullReferenceException("target 为 null");
            }
            if (datas == null)
            {
                throw new ArgumentNullException("datas");
            }
            var ldv2_2 = target.DataValidation.AddListDataValidation();
            for (int i = 0; i < datas.Length; i++)
            {
                ldv2_2.Formula.Values.Add(datas[i]);
            }
        }
    }
}
