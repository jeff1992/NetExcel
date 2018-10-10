﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetExcel;

namespace ExportTest
{
	class Program
	{
		static void Main(string[] args)
        {
            Random random = new Random();
            Dictionary<string, IEnumerable<string>> dic = new Dictionary<string, IEnumerable<string>>();
            dic.Add("Fruit", new string[] { "Peach", "Plum", "Banana", "Pear" });
            dic.Add("Vegetable", new string[] { "Cabbage", "Potato", "Cucumber", "Bear" });

            //构造model
            var order = new
            {
                ProjectName = "Gray wolf's birthday party",
                Name = "Jeff",
                CreatedAt = DateTime.Now,
                BuyerName = "Bill",
                Cates = dic.Select(m => new
                {
                    Name = m.Key,
                    Items = m.Value.Select(n => new
                    {
                        Name = n,
                        Price = (decimal)random.Next(1, 100),
                        Amount = random.Next(1, 100)
                    }).ToList()
                })
            };
            ExcelTpl render = new ExcelTpl("tpl.xlsx");
            render.KeyValues.Add("order", order);
            var fileName = DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            render.RenderAndSave(fileName);
            System.Diagnostics.Process.Start(fileName);
        }
	}
}
