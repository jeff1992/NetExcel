using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeOpenXml.FormulaParsing;
using NCalc.Domain;

namespace NetExcel
{
    internal class ExlInterpreter
    {
        ExcelWorksheet sheet;
        List<ExcelRange> mergedRegions;

        int rows = 0;
        int columns = 0;
        //当前编译的行
        int interpRow = 1;
        //当前编译的字符位置
        int interpChar = 0;
        int outputRow = 1;

        Regex regReplace = new Regex(@"{\S+?}");
        Regex regExp = new Regex(@"([a-z]+)\s*\([\S|\s]+?\)\s*");
        Regex regBlockStart = new Regex(@"^{\s*");
        Regex regBlockEnd = new Regex(@"^}\s*");
        Regex regDisplay = new Regex(@"(^{\s*$|^{\s*}\s*|^(\s*)})");
        Regex regMethod = new Regex(@"([a-zA-Z|\.]+)\s*\([\S|\s]+?\)\s*");
        Regex regParam = new Regex(@"[a-zA-Z][a-zA-Z\d.]+");
        public ExlInterpreter(OfficeOpenXml.ExcelWorksheet sheet)
        {
            this.sheet = sheet;
            this.rows = sheet.Dimension.End.Row;
            this.columns = sheet.Dimension.End.Column;
            mergedRegions = new List<ExcelRange>();
        }
        public void Complie(Dictionary<string, object> values)
        {
            this.outputRow = this.rows + 1;
            this.Run(values);
            foreach (var m in mergedRegions)
            {
                sheet.Cells[m.Address].Merge = true;
            }

            sheet.DeleteColumn(1);
            //删除一行后，下面一行变成第一行。批量删除的方法有问题，所以一行一行删除
            for (var i = this.rows; i >=1 ; i--)
            	sheet.DeleteRow(i);
        }
        //以block为递归单元
        void Run(Dictionary<string, object> data)
        {
            while (this.interpRow <= this.rows)
            {
                var cellStr = sheet.Cells[this.interpRow, 1].Text.Trim();
                if (string.IsNullOrWhiteSpace(cellStr))
                {
                    this.SetRow(this.interpRow, this.outputRow++, data);
                    this.interpRow++;
                }
                else
                {
                    var str = cellStr.Substring(this.interpChar, cellStr.Length - this.interpChar);
                    if (regExp.IsMatch(str))
                    {
                        var exp = regExp.Match(str).Value;
                        this.interpChar += exp.Length;
                        var leftIndex = exp.IndexOf('(');
                        var rightIndex = exp.LastIndexOf(')');
                        var expName = exp.Substring(0, leftIndex).Trim(); //获取语句类型
                        var expParam = exp.Substring(leftIndex + 1, rightIndex - leftIndex - 1);  //获取括号中的内容
                        switch (expName)
                        {
                            case "for":
                                var arr = expParam.Split(new string[] { " in " }, 2, StringSplitOptions.RemoveEmptyEntries).Select(m => m.Trim());
                                if (arr.Count() != 2)
                                    throw new Exception("Invalid expression：" + expParam);
                                var pre = arr.First().Split(',');
                                var key = pre.First();
                                var indexName = pre.ElementAtOrDefault(1);
                                var value = GetValue(arr.ElementAt(1), data) as IEnumerable;
                                if (value == null)
                                    throw new Exception($"{arr.ElementAt(1)} can not be Enumerated");
                                var startRow = this.interpRow;
                                var startChar = this.interpChar;
                                var index = 1;

                                foreach (var obj in value)
                                {
                                    data.Add(key, obj);
                                    if (!string.IsNullOrWhiteSpace(indexName))
                                        data.Add(indexName, index);
                                    this.interpRow = startRow;
                                    this.interpChar = startChar;
                                    this.Run(data);
                                    data.Remove(key);
                                    if (!string.IsNullOrWhiteSpace(indexName))
                                        data.Remove(indexName);
                                    index++;
                                }
                                // 当集合数据为空时，跳过当前循环
                                if (index == 1)
                                {
                                    this.interpRow++;
                                    this.interpChar = 0;
                                }
                                break;
                            default:
                                throw new Exception($"Unknown expression: {expName}");
                        }
                        this.Run(data);
                    }
                    else if (regBlockStart.IsMatch(str))
                    {
                        if (regDisplay.IsMatch(str))
                        {
                            this.SetRow(this.interpRow, this.outputRow++, data);
                        }
                        this.interpChar += regBlockStart.Match(str).Value.Length;
                    }
                    else if (regBlockEnd.IsMatch(str))
                    {
                        if (regDisplay.IsMatch(cellStr))
                        {
                            this.SetRow(this.interpRow, this.outputRow++, data);
                        }
                        this.interpChar += regBlockEnd.Match(str).Value.Length;
                        break;
                    }
                    else if (string.IsNullOrWhiteSpace(str))
                    {
                        this.interpChar = 0;
                        this.interpRow++;
                    }
                    else
                    {
                        throw new Exception($"Invalid expression: {str}");
                    }
                }
            }
        }
        void SetRow(int fromRow, int newRow, Dictionary<string, object> data)
        {
            Console.WriteLine($"Rending line {fromRow} to line {newRow}");
            sheet.InsertRow(newRow, 1, fromRow);
            sheet.Row(newRow).Height = sheet.Row(fromRow).Height;
            for (var i = 2; i <= columns; i++)
            {
                Console.WriteLine($"Cell {((char)('A' + i)).ToString()}{fromRow}");
                var mer = sheet.MergedCells[fromRow, i];
                //这里可以加历史记录，判断是否已经合并过，防止重复合并
                if (mer != null)
                {
                    ExcelAddress addr = new ExcelAddress(mer);
                    var merCel = sheet.Cells[addr.Start.Row + (newRow - fromRow), addr.Start.Column, addr.End.Row + (newRow - fromRow), addr.End.Column];
                    if (!merCel.Merge)
                    {
                        Console.WriteLine($"Merge cell {merCel.ToString()}");
                        // 判断单元格是否已经包含在需要合并单元格列表中
                        if (!mergedRegions.Any(a => a.Address == merCel.Address))
                        {
                            mergedRegions.Add(merCel);
                        }

                        //try
                        //{
                        //  merCel.Merge = true;
                        //}
                        //catch (Exception e)
                        //{
                        //}
                    }
                }
                var fromCell = sheet.Cells[fromRow, i];
                var newCell = sheet.Cells[newRow, i];
                if (!string.IsNullOrWhiteSpace(fromCell.FormulaR1C1))
                {
                    newCell.FormulaR1C1 = fromCell.FormulaR1C1;
                }
                else
                {
                    var org = fromCell.Value;
                    if (org != null)
                    {
                        var orgStr = org.ToString();
                        if (orgStr.StartsWith("#="))  //公式
                        {
                            newCell.FormulaR1C1 = ReplaceParam(orgStr.Substring(1), data).ToString();
                        }
                        else
                        {
                            newCell.Value = ReplaceParam(org.ToString(), data);
                        }
                    }
                }
            }
        }
        object ReplaceParam(string name, Dictionary<string, object> data)
        {
            var matched = regReplace.Matches(name);
            var keys = matched.Cast<Match>().Select(m => m.Value).Distinct();
            if (keys.Count() == 1 && keys.First().Length == name.Length)
            {
                var key = keys.First();
                return ExecExpression(key.Substring(1, key.Length - 2), data);
            }
            else
            {
                foreach (var key in keys)
                {
                    var param = key.Substring(1, key.Length - 2);
                    var val = ExecExpression(param, data);
                    name = name.Replace(key, val.ToString());
                }
                return name;
            }
        }


        object ExecExpression(string expression, Dictionary<string, object> data)
        {
            Dictionary<string, object> paramenters = new Dictionary<string, object>();
            var key = 'a';
            var newExp = regParam.Replace(expression, (Match x) =>
            {
                paramenters.Add(key.ToString(), GetValue(x.Value, data));
                return key++.ToString();
            });
            var e = new NCalc.Expression(newExp);
            e.Parameters = paramenters;
            return e.Evaluate();
                //			var list = expression.Split('+', '-', '*', '/', '%');
                //			if (list.Length > 1)
                //			{
                //
                //				return null;
                //			}
                //			else
                //			{

                return GetValue(expression, data);
            //			}
        }

        object GetValue(string name, Dictionary<string, object> data)
        {
            object value = null;
            var arr = name.Split('.');
            try
            {
                value = data[arr.First()];
            }
            catch
            {
                throw new Exception("Variable not found：" + arr.First());
            }
            foreach (var pName in arr.Skip(1))
            {
                var prop = value.GetType().GetProperty(pName);
                if (prop == null)
                {
                    throw new Exception($"Type \"{value.GetType().Name}\" does not contains member: {pName}");
                }
                value = prop.GetValue(value, null);
            }
            return value;
        }
    }
}
