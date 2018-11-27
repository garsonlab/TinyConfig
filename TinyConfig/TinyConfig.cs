using Aspose.Cells;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;

namespace TinyConfig
{
    public class Converter
    {
        #region 解析选项
        public struct Options
        {
            public string excelPath;//excel
            public string serverFolder;//server csv
            public string clientFolder;//client csv
            public string csFolder;//csharp
            public string nameSpace;//csharp 命名空间
        }

        private readonly Options options;

        public Converter(Options options)
        {
            this.options = options;
            if (!File.Exists(options.excelPath))
                throw new NullReferenceException($"No File : {options.excelPath}");

            try
            {
                Workbook workbook = new Workbook(options.excelPath, new LoadOptions());
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    ParseSheet(worksheet);
                }
        }
            catch (Exception e)
            {
                throw e;
            }
}
        #endregion

        private void ParseSheet(Worksheet worksheet)
        {
            Cells cells = worksheet.Cells;
            string outputName = GetCell(cells, 0, 1);//output
            if(string.IsNullOrEmpty(outputName))
                throw new Exception("Output Name is Null in Row:0, Col:1");
            
            ExportCSV(cells, 1, options.serverFolder, outputName);//server
            ExportCSV(cells, 2, options.clientFolder, outputName);//client
            ExportCS(cells, 2, 3, options.csFolder, outputName, options.nameSpace);//client cs
        }

        /// <summary>
        /// 转换CSV
        /// </summary>
        /// <param name="cells">当前Sheet的所有格子</param>
        /// <param name="idx">当前用来转换的标志行</param>
        /// <param name="directory">转换文件路径</param>
        /// <param name="title">转换文件名</param>
        private void ExportCSV(Cells cells, int idx, string directory, string title)
        {
            if(!CheckDirectory(directory, out string path))//当前路径为空，不转换
                return;
            path = $"{path}/{title}.csv";

            Row flag = cells.GetRow(idx);
            StringBuilder builder = new StringBuilder();

            bool exprot = ExportCheck(flag, builder, out List<bool> exports);
            if(!exprot)
                return;

            for (int i = 5; i < cells.Rows.Count; i++)
            {
                Row row = cells.GetRow(i);
                if(row == null)
                    break;
                for (int j = 1; j < exports.Count; j++)
                {
                    if(!exports[j])
                        continue;

                    string v = GetCell(row, j);
                    builder.Append(v);
                    builder.Append(",");
                }

                builder.Remove(builder.Length - 1, 1);
                builder.AppendLine();
            }

            File.WriteAllText(path, builder.ToString(), new UTF8Encoding(false));
        }

        /// <summary>
        /// 转换CS文件
        /// </summary>
        /// <param name="cells">当前Sheet的所有格子</param>
        /// <param name="idx">当前用来转换的标志行</param>
        /// <param name="sidx">类型标志行</param>
        /// <param name="directory">转换文件路径</param>
        /// <param name="title">转换文件名、类名</param>
        private void ExportCS(Cells cells, int idx, int sidx, string directory, string title, string nameSpace)
        {
            if (!CheckDirectory(directory, out string path))//当前路径为空，不转换
                return;
            path = $"{path}/{title}.cs";

            Row row = cells.GetRow(idx);
            Row type = cells.GetRow(sidx);
            
            if(!ExportCheck(row, null, out List<bool> exports))
                return;
            
            TableFlags main = new TableFlags() { ismain = true, name = title };
            Dictionary<string, TableFlags> flagDic = new Dictionary<string, TableFlags>() { {title, main } };

            for (int i = 1; i < exports.Count; i++)
            {
                if(!exports[i])
                    continue;

                string name = GetCell(row, i);
                string symbol = GetCell(type, i);
                
                ParseNode(name, symbol, main, flagDic);
            }

            List<string> finder = ParseFinder(row, type, exports);

            StringBuilder builder = new StringBuilder();
            builder.AppendLine("using System.Collections.Generic;\n");

            string tab = "";
            bool useSpace = false;
            if (!string.IsNullOrEmpty(nameSpace))
            {
                useSpace = true;
                builder.AppendLine($"namespace {nameSpace}\n{{");
                tab = "\t";
            }

            List<TableFlags> flags = flagDic.Values.ToList();
            while (flags.Count > 0)
            {
                builder.AppendLine(GenegrateCode(main, tab, flagDic));

                flags.Remove(main);
                if (flags.Count > 0)
                {
                    main = flags[0];
                }
            }

            builder.AppendLine(GenegrateHelper(title, tab, row, exports, finder));
            

            if (useSpace)
                builder.AppendLine("\t}\n}");
            else
                builder.AppendLine("}");

            File.WriteAllText(path, builder.ToString(), new UTF8Encoding(false));
        }

        private void ParseNode(string name, string symbol, TableFlags table, Dictionary<string, TableFlags> dic)
        {
            if (!name.Contains("#"))//单个
            {
                table.parms.Add(name, GetConvertSymbol(symbol));
            }
            else
            {
                string[] parts = name.Split('#');
                List<uint> size = new List<uint>();
                int n0 = 1;
                for (; n0 < parts.Length; n0++)
                {
                    if (uint.TryParse(parts[n0], out uint v))
                    {
                        size.Add(v);
                    }
                    else
                    {
                        break;
                    }
                }

                if (size.Count > 0)//表示当前是数组
                {
                    string s;
                    if(n0 >= parts.Length)
                        s = GetConvertSymbol(symbol);
                    else
                    {
                        s = GetTypeUpper(parts[0]);//表示当前是类或结构体
                        s = GetConvertSymbol(s);//转换成类
                    }


                    if (table.array.TryGetValue(parts[0], out ArrayFlags array))
                    {
                        //找出当前数组各维度最大值
                        List<uint> compare = array.deps;

                        if (compare.Count >= size.Count)
                        {
                            for (int n1 = 0; n1 < size.Count; n1++)
                            {
                                if (compare[n1] < size[n1])
                                    compare[n1] = size[n1];
                            }
                            array.deps = compare;
                        }
                        else
                        {
                            for (int n1 = 0; n1 < compare.Count; n1++)
                            {
                                if (size[n1] < compare[n1])
                                    size[n1] = compare[n1];
                            }
                            array.deps = size;
                        }
                    }
                    else
                    {
                        table.array.Add(parts[0], new ArrayFlags()
                        {
                            deps = size,
                            name = parts[0],
                            symbol = s
                        });
                    }
                        
                    if(n0 < parts.Length)//说明不是完整的数组，后面是类成员
                    {
                        if (!dic.ContainsKey(s))
                        {
                            TableFlags tf = new TableFlags()
                            {
                                name = s
                            };
                            dic.Add(s, tf);

                            string f = parts[n0];
                            for (int i = n0+1; i < parts.Length; i++)
                            {
                                f += "#" + parts[i];
                            }
                            ParseNode(f, symbol, tf, dic);
                        }
                    }
                }
                else
                {
                    string s = GetTypeUpper(parts[0]);
                    s = GetConvertSymbol(s);

                    if(!table.parms.ContainsKey(parts[0]))//加入到变量
                        table.parms.Add(parts[0], s);
                    if (!dic.TryGetValue(s, out TableFlags tf))
                    {
                        tf = new TableFlags()
                        {
                            name = s
                        };
                        dic.Add(s, tf);
                    }
                    string f = parts[1];
                    for (int i = 2; i < parts.Length; i++)
                    {
                        f += "#" + parts[i];
                    }
                    ParseNode(f, symbol, tf, dic);
                }
            }
        }

        private List<string> ParseFinder(Row flag, Row symbol, List<bool> exports)
        {
            List<string> finder = new List<string>();
            for (int i = 1; i < exports.Count; i++)
            {
                if(!exports[i])
                    continue;

                if(GetCell(symbol, i) == "_key")
                    finder.Add(GetCell(flag, i));
            }

            return finder;
        }

        /// <summary>转换类型</summary>
        private string GetConvertSymbol(string symbol)
        {
            switch (symbol)
            {
                case "text":
                    return "string";
                case "_key":
                    return "int";
                case "byte":
                case "int":
                case "long":
                case "float":
                case "uint":
                case "double":
                    return symbol;
                default:
                    return GetTypeUpper(symbol) + "Class";
            }
        }

        /// <summary>首字母大写</summary>
        private string GetTypeUpper(string symbol)
        {
            return symbol.Substring(0, 1).ToUpper() + symbol.Substring(1);
        }

        /// <summary>首字母小写</summary>
        private string GetNameLower(string name)
        {
            return name.Substring(0, 1).ToLower() + name.Substring(1);
        }

        /// <summary>生成Class</summary>
        private string GenegrateCode(TableFlags flag, string tab, Dictionary<string, TableFlags> dic)
        {
            tab = flag.ismain ? tab : tab+"\t";
            StringBuilder builder = new StringBuilder();
            builder.AppendLine($"{tab}public class {flag.name}");
            builder.AppendLine($"{tab}{{");
            
            bool needConstructor = false;
            StringBuilder constructor = new StringBuilder();

            foreach (var parm in flag.parms)
            {
                builder.AppendLine($"{tab}\tpublic {parm.Value} {parm.Key};");

                if (dic.ContainsKey(parm.Value))
                {
                    needConstructor = true;
                    if (constructor.Length > 0)
                        constructor.AppendLine();
                    constructor.Append($"{tab}\t\t{parm.Key} = new {parm.Value}();");
                }
            }

            foreach (var array in flag.array)
            {
                var deps = array.Value.deps;
                string depc = "[";
                string depn = $"[{deps[0]}";
                for (int i = 1; i < deps.Count; i++)
                {
                    depc += ", ";
                    depn += $", {deps[i]}";
                }
                depn += "]";
                depc += "]";

                builder.AppendLine($"{tab}\tpublic {array.Value.symbol}{depc} {array.Value.name};");
                needConstructor = true;

                if (constructor.Length > 0)
                    constructor.AppendLine();
                constructor.Append($"{tab}\t\t{array.Value.name} = new {array.Value.symbol}{depn};");

                //判断class数组
                if (dic.ContainsKey(array.Value.symbol))
                {
                    string ct = $"{tab}\t\t";
                    string cp = "";
                    for (int i = 0; i < deps.Count; i++)
                    {
                        constructor.AppendLine();
                        constructor.Append($"{ct}for(int c{i} = 0; c{i} < {deps[i]}; c{i}++)");
                        ct += "\t";

                        if (i == 0)
                            cp = "[c0";
                        else
                            cp += $", c{i}";
                    }

                    constructor.AppendLine();
                    constructor.Append($"{ct}{array.Value.name}{cp}] = new {array.Value.symbol}();");
                }
            }

            if (needConstructor)
            {
                builder.AppendLine();
                builder.AppendLine($"{tab}\tpublic {flag.name}()");
                builder.AppendLine($"{tab}\t{{");
                builder.AppendLine(constructor.ToString());
                builder.AppendLine($"{tab}\t}}");
            }

            if (!flag.ismain)
                builder.AppendLine($"{tab}}}");
            return builder.ToString();
        }

        /// <summary>生成解析</summary>
        private string GenegrateHelper(string title, string tab, Row row, List<bool> exports, List<string> finder)
        {
            StringBuilder builder = new StringBuilder();

            builder.AppendLine($"{tab}\tprivate static Table.Index _index_ = new Table.Index();");
            builder.AppendLine($"{tab}\tprivate static List<{title}> _values_ = new List<{title}>();");
            builder.AppendLine($"{tab}\tpublic static List<{title}> Values => _values_;\n");
            builder.AppendLine($"{tab}\tpublic static bool Load(string str)");
            builder.AppendLine($"{tab}\t{{");
            builder.AppendLine($"{tab}\t\tTable.Parser parser = new Table.Parser(new Table.Parser.Options {{space = ',',title = 0, context = 1}});");
            builder.AppendLine($"{tab}\t\treturn parser.Load(str, new Handler());");
            builder.AppendLine($"{tab}\t}}\n");

            builder.AppendLine($"{tab}\tpublic static {title} Find({GenegrateParm(finder)})");
            builder.AppendLine($"{tab}\t{{");
            builder.AppendLine($"{tab}\t\tvar key = _index_.Get();");
            builder.AppendLine($"{tab}\t\tkey.Clear();");
            foreach (var par in finder)
            {
                builder.AppendLine($"{tab}\t\tkey.Add({par}.ToString());");
            }
            builder.AppendLine($"{tab}\t\tint idx = _index_.Of(key.ToString());");
            builder.AppendLine($"{tab}\t\tif (idx < 0) return null;");
            builder.AppendLine($"{tab}\t\treturn _values_[idx];");
            builder.AppendLine($"{tab}\t}}\n");

            builder.AppendLine($"{tab}\tprivate class Handler : Table.IHandler");
            builder.AppendLine($"{tab}\t{{");
            builder.AppendLine($"{tab}\t\tprivate {title} def;");
            builder.AppendLine($"{tab}\t\tvoid Table.IHandler.Start()");
            builder.AppendLine($"{tab}\t\t{{");
            builder.AppendLine($"{tab}\t\t\tdef = new {title}();");
            builder.AppendLine($"{tab}\t\t}}");

            builder.AppendLine($"{tab}\t\tvoid Table.IHandler.End()");
            builder.AppendLine($"{tab}\t\t{{");
            builder.AppendLine($"{tab}\t\t\t_values_.Add(def);");
            builder.AppendLine($"{tab}\t\t\tvar key = _index_.Get();");
            builder.AppendLine($"{tab}\t\t\tkey.Clear();");
            foreach (var par in finder)
            {
                builder.AppendLine($"{tab}\t\t\tkey.Add(def.{par}.ToString());");
            }
            builder.AppendLine($"{tab}\t\t\t_index_.Add(key.ToString(), _values_.Count-1);");
            builder.AppendLine($"{tab}\t\t}}");

            builder.AppendLine($"{tab}\t\tbool Table.IHandler.Value(string key, string value)");
            builder.AppendLine($"{tab}\t\t{{");
            builder.AppendLine($"{tab}\t\t\tswitch(key)");
            builder.AppendLine($"{tab}\t\t\t{{");
            for (int i = 0; i < exports.Count; i++)
            {
                if(!exports[i])
                    continue;
                string v = GetCell(row, i);
                builder.AppendLine(GenegrateValue(tab + "\t\t\t\t", v));
            }
            builder.AppendLine($"{tab}\t\t\t}}");
            builder.AppendLine($"{tab}\t\t\treturn false;");
            builder.AppendLine($"{tab}\t\t}}");
            builder.AppendLine($"{tab}\t}}");
            return builder.ToString();
        }

        private string GenegrateParm(List<string> finder)
        {
            if(finder.Count <= 0)
                throw new Exception("No ‘_key’in List");

            string parm = "int " + finder[0];
            for (int i = 1; i < finder.Count; i++)
            {
                parm += ", int " + finder[i];
            }

            return parm;
        }

        /// <summary>生成value case</summary>
        private string GenegrateValue(string tab, string v)
        {
            string[] parts = v.Split('#');
            string parm = parts[0];
            List<uint> dep = new List<uint>();
            int n0 = 1;

            while (n0 < parts.Length)
            {
                if (uint.TryParse(parts[n0], out uint a))
                {
                    dep.Add(a);
                }
                else
                {
                    if (dep.Count > 0)
                    {
                        parm += "[" + (dep[0] - 1);
                        for (int i = 1; i < dep.Count; i++)
                        {
                            parm += ", " + (dep[i]-1);
                        }
                        parm += "]";
                    }
                    //else
                    //{
                        parm += "." + parts[n0];
                    //}
                    dep.Clear();
                }
                ++n0;
            }
            if (dep.Count > 0)
            {
                parm += "[" + (dep[0] - 1);
                for (int i = 1; i < dep.Count; i++)
                {
                    parm += ", " + (dep[i] - 1);
                }
                parm += "]";
            }
            
            return $"{tab}case \"{v}\":\n{tab}\t return StringUtil.Parse(value, out def.{parm});";
        }

        class TableFlags
        {
            public bool ismain;
            public string name;
            public Dictionary<string, string> parms = new Dictionary<string, string>();//单个变量
            public Dictionary<string, ArrayFlags> array = new Dictionary<string, ArrayFlags>();
        }

        class ArrayFlags
        {
            public string name;
            public string symbol;
            public List<uint> deps = new List<uint>();
        }
        
        private bool ExportCheck(Row flag, StringBuilder builder, out List<bool> export)
        {
            export = new List<bool>(){false};

            int col = flag.LastCell.Column;
            if (col < 1 || string.IsNullOrEmpty(GetCell(flag, col)))
                return false;

            bool effective = false;
            for (int i = 1; i <= col; i++)
            {
                string title = GetCell(flag, i);
                if (string.IsNullOrEmpty(title))
                {
                    export?.Add(false);
                }
                else
                {
                    effective = true;
                    export?.Add(true);
                    builder?.Append(title);
                    builder?.Append(",");
                }
            }

            builder?.Remove(builder.Length - 1, 1);
            builder?.AppendLine();
            return effective;
        }

        private string GetCell(Row row, int idx)
        {
            Cell cell = row[idx];
            return cell?.DisplayStringValue.Trim();
        }

        private string GetCell(Cells cells, int row, int col)
        {
            Cell cell = cells.GetCell(row, col);
            return cell?.DisplayStringValue.Trim();
        }

        private bool CheckDirectory(string path, out string dir)
        {
            dir = null;
            if (string.IsNullOrEmpty(path))
                return false;

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            dir = path;
            return true;
        }
    }
}
