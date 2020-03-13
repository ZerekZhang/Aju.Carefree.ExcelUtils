using Aju.Carefree.ExcelUtils;
using System;
using System.Collections.Generic;

namespace Aju.Carefree
{
    class Program
    {
        static void Main(string[] args)
        {
            var list = new List<Person>();
            for (int i = 0; i < 10; i++)
            {
                list.Add(new Person
                {
                    Sex = "Sex" + i,
                    Age = "Age" + i,
                    Brithday = DateTime.Now,
                    Name = "Name" + i
                });
            }
            EPPlusToExcelHelper.CreateExcelByList(@"D:\123.xlsx", "", list, ws =>
            {
                //ws.Cells[1, 1, 1, 2].Merge = true;
                //ws.Cells[1, 1, 1, 2].Merge = true;
                //ws.Cells[1, 1].Value = "姓名";

                //ws.Cells[1, 2].Value = "年龄";

                //ws.Cells[1, 3, 1, 4].Merge = true;
                //ws.Cells[1, 3].Value = "生日";

                //ws.Cells[1, 4].Value = "性别";
                //return (ws, 1);
            });
            Console.WriteLine("Hello World!");
        }
    }

    public class Person
    {
        [ExcelColumn(1, "姓名")]
        public string Name { get; set; }
        [ExcelColumn(2, "姓名")]
        public string Age { get; set; }
        [ExcelColumn(4, "姓名")]
        public string Sex { get; set; }
        [ExcelColumn(3, "姓名")]
        public DateTime Brithday { get; set; }
    }
}
