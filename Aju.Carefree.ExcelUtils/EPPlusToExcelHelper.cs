using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Aju.Carefree.ExcelUtils
{
    /// <summary>
    /// EPPlusToExcel 
    /// <para>EPPlus 导出 EXCEL 帮助类</para>
    /// </summary>
    public static class EPPlusToExcelHelper
    {
        /// <summary>
        /// 创建Excel 并保存 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath">excel保持路径</param>
        /// <param name="sheetName">sheet 名称</param>
        /// <param name="listData">数据</param>
        /// <param name="excelHader">表头自定义，如果为空，则自动生成表头，复杂表头请自定义
        /// <para>int：表格内容从第几行开始</para>  
        /// </param>
        /// <param name="excelStyle">excel 样式</param>
        public static void CreateExcelByList<T>(string filePath, string sheetName, List<T> listData,
            Func<ExcelWorksheet, (ExcelWorksheet, int)> excelHader = null, Func<ExcelWorksheet, ExcelWorksheet> excelStyle = null) where T : class
        {
            Type t = typeof(T);
            if (string.IsNullOrEmpty(filePath)) throw new NullReferenceException($"参数{nameof(filePath)}不能为空.");
            if (string.IsNullOrEmpty(sheetName)) sheetName = "sheet1";
            if (listData == null || listData.Count == 0) throw new NullReferenceException("导出内容不能为空.");
            if (File.Exists(filePath))
                File.Delete(filePath);
            FileInfo newfile = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(newfile))
            {
                //在工作簿中获得第一个工作表
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
                //表头
                int row = 1;
                if (excelHader == null)
                {
                    t.GetProperties().ToList().ForEach((n) =>
                    {
                        if (!n.IsDefined(typeof(ExcelColumnAttribute), true))
                            return;
                        ExcelColumnAttribute attribute = (ExcelColumnAttribute)n.GetCustomAttribute(typeof(ExcelColumnAttribute), true);
                        if (attribute != null)
                            worksheet.Cells[row, attribute.Column].Value = attribute.Name;
                    });
                }
                else
                    (worksheet, row) = excelHader.Invoke(worksheet);
                foreach (var item in listData)
                {
                    foreach (var p in t.GetProperties())
                    {
                        if (p.IsDefined(typeof(ExcelColumnAttribute), true))
                        {
                            ExcelColumnAttribute attribute = (ExcelColumnAttribute)p.GetCustomAttribute(typeof(ExcelColumnAttribute), true);
                            if (attribute != null)
                            {
                                var value = p.GetValue(item);
                                if (value != null && !string.IsNullOrEmpty(value.ToString()))
                                {
                                    worksheet.Cells[row + 1, attribute.Column].Value = value.ToString();
                                }
                            }
                        }
                    }
                    row++;
                }
                if (excelStyle != null)
                    worksheet = excelStyle.Invoke(worksheet);
                package.Save();
            }
        }
    }
}
