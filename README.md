## Introduction
Here you can simply export a excel using excel template like html template. You will never need to set cell styles in your code. Instead, all the cell styles and expressions can be set in the excel visibly. You will only need to code like this:
```c#
var tpl = new ExcelTemplate("tpl.xls");
tpl.Values.Add("order", order);
tpl.SaveAs("newfile.xls");
```

## Dependencis
```bash
.Net Framework ≥ 4.5
```
OR
```bash
.Net Core ≥ 2.0
```

## Install
```nuget
Install-Package NetExcel
```

## Usage
--First, make your template
<a href="https://github.com/jeff1992/NetExcel/blob/master/tpl.png">
	<img src="https://github.com/jeff1992/NetExcel/blob/master/tpl.png">
</a>
1. Control expression<br>
	Please keep the first column to write control expression like "for(...)"<br>
	Supports:<br>
		for(item in items)<br>
		for(item,index in items)	#index will count up from 1<br>
2. Value display<br>
	{user.name}<br>
	note: method or operation not supported now<br>
	
3. Make it work in your code

```c#
using System;
using System.Collections.Generic;
using System.Linq;
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
			var tpl = new ExcelTemplate("tpl.xlsx");
			tpl.KeyValues.Add("order", order);
			var fileName = DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

			//bellow is the main method
			tpl.SaveAs(fileName);
			//open file
			System.Diagnostics.Process.Start(fileName);
        }
	}
}

```

## License

[MIT](https://github.com/jeff1992/NetExcel/blob/master/LICENSE)

Copyright (c) 2017-present Jeff.Wang
