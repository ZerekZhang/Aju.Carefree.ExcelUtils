using System;

namespace Aju.Carefree.ExcelUtils
{
    /// <summary>
    /// Excel 列
    /// </summary>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        ///  Excel 列数
        /// </summary>
        public int Column { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="column">列数 第几列 </param>
        /// <param name="name">标题名称</param>
        public ExcelColumnAttribute(int column, string name)
        {
            Column = column;
            Name = name;
        }
    }
}
