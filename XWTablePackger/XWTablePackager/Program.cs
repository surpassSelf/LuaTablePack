using System;
using Spire.Xls;
using System.IO;
using System.Text;
using System.Collections.Generic;
//using System.Windows.Forms;
using System.Threading;

namespace TablePacker
{
    class Program
    {
        static long timeKeeping = 0;
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Black;
            Console.BackgroundColor = ConsoleColor.White;
            Console.Clear();

            Console.SetWindowSize(130, 40);
            Console.BufferHeight = short.MaxValue - 1;

            timeKeeping = DateTime.Now.Ticks;

            //args = new string[] { @"F:\3DGuoZhan\Client\Branch\Excel\Skill-技能配置表.xlsx", @"F:\3DGuoZhan\Client\Branch\3DClient\Assets\Lua\Config\data" };

            if (args.Length == 0)
            {
                PrintError("Please input file path");
                End();
                return;
            }

            string excelFile = args[0].Replace('/', '\\');

            string outputDir = "";
            if (args.Length > 1)
            {
                outputDir = args[1].Replace('/', '\\').TrimEnd('\\');
            }
            else
            {
                outputDir = File.Exists(excelFile) ? Path.GetDirectoryName(excelFile) : excelFile;
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            List<Thread> threadList = new List<Thread>();

            if (File.Exists(excelFile))
            {
                TableHandler.Build(excelFile, outputDir, threadList);
            }
            else if (Directory.Exists(excelFile))
            {
                string[] files = Directory.GetFiles(excelFile, "*.xlsx");
                if (files.Length == 0)
                {
                    PrintError("Don't find xlsx format file in: " + excelFile);
                }
                else
                {
                    int count = files.Length;
                    for (int i = 0; i < count; i++)
                    {
                        Console.Write(string.Format("[{0}/{1}] ", (i + 1).ToString().PadLeft(count.ToString().Length, ' '), count));
                        TableHandler.Build(files[i], outputDir, threadList);
                    }
                }
            }
            else
            {
                PrintError("Don't find file or directory: " + excelFile);
            }

            while (threadList.Count > 0)
            {
                Thread.Sleep(1000);
            }

            End();
        }


        class TableHandler
        {
            StringBuilder tempSb = new StringBuilder();

            public static void Build(string file, string outputDir, List<Thread> threadList)
            {
                if (file.Contains("~$") || file.Contains(".~"))
                {
                    Console.Write(Path.GetFullPath(file) + " >> ");
                    PrintGreenInfo("跳过");
                    return;
                }

                string error;
                while (IsFileInUse(file, out error))
                {
                    PrintError(error + ": " + file);
                    Console.ReadLine();
                    //if (MessageBox.Show("跳过此文件点取消，关闭文件后重试点确定！", "", MessageBoxButtons.RetryCancel) == DialogResult.Cancel)
                    //    return;
                    //Console.WriteLine("跳过此文件按Esc键，关闭文件后重试按Enter回车键！");
                    //var consoleKey = Console.ReadKey();
                    //while (consoleKey.Key != ConsoleKey.Enter)
                    //{
                    //    if (consoleKey.Key == ConsoleKey.Escape)
                    //        return;
                    //}
                }

                string extension = Path.GetExtension(file);

                if (extension != ".xlsx")
                {
                    PrintError("The tool only support xlsx format file: " + file);
                    return;
                }

                Workbook workbook = new Workbook();
                workbook.LoadFromFile(file);

                Worksheet sheet = workbook.Worksheets[0];

                if (sheet.LastRow > 5000)
                {
                    Thread thread = new Thread(p =>
                    {
                        new TableHandler().BuildTable(file, outputDir, sheet);
                        threadList.Remove((Thread)p);
                    });
                    threadList.Add(thread);
                    thread.Start(thread);
                }
                else
                    new TableHandler().BuildTable(file, outputDir, sheet);
            }

            string fileName;

            void BuildTable(string file, string outputDir, Worksheet sheet)
            {
                CellRange range = sheet.Range[sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn];

                string sheetName = sheet.Name;

                //------------------------------------------------------------------

                fileName = Path.GetFileName(file);
                Console.Write(Path.GetFullPath(file) + " >> ");
                string outputFile = outputDir + "\\" + sheetName + ".lua";
                Program.PrintGreenInfo(outputFile);

                //------------------------------------------------------------------

                if (sheet.IsEmpty)
                {
                    PrintError("Sheet is empty: " + fileName);
                    return;
                }
                //PrintCellRange(range);

                List<TableFieldInfo> fieldInfos = new List<TableFieldInfo>();
                for (int i = 1; i <= range.LastColumn; i++)
                {
                    TableFieldType t;
                    string name = range[3, i].DisplayedText;
                    if (Enum.TryParse<TableFieldType>(range[2, i].DisplayedText, out t) && !string.IsNullOrEmpty(name))
                    {
                        fieldInfos.Add(new TableFieldInfo { name = name, type = t, column = i });
                    }
                }

                bool hasDefault = false;
                Dictionary<string, int> countDic = new Dictionary<string, int>();
                string key;
                int maxCount = 0;
                for (int column = 0; column < fieldInfos.Count; column++)
                {
                    countDic.Clear();
                    for (int row = 4; row <= range.LastRow; row++)
                    {
                        key = range[row, fieldInfos[column].column].DisplayedText;

                        int count;
                        countDic.TryGetValue(key, out count);
                        count = countDic[key] = count + 1;

                        if (count > 3 && count > maxCount)
                        {
                            fieldInfos[column].hasDefault = true;
                            fieldInfos[column].defaultV = key;
                            hasDefault = true;
                        }
                    }
                }

                if (fieldInfos.Count == 0)
                {
                    PrintError("Don't find client field: " + fileName);
                    return;
                }

                //------------------------------------------------------------------

                StringBuilder sb = new StringBuilder();
                sb.Append("-- 此文件工具自动生成，不要修改\n\n");

                sb.Append("local ").Append(sheetName).Append(" =");

                sb.Append("\n{\n");

                int firstFieldColumn = fieldInfos[0].column;
                string content;
                //uint v;
                List<string> idList = new List<string>();

                for (int row = 4; row <= range.LastRow; row++)
                {
                    content = range[row, firstFieldColumn].DisplayedText;

                    if (idList.Contains(content))
                    {
                        PrintDataException(row, firstFieldColumn, fileName + " Has the same id: " + content);
                        continue;
                    }
                    else
                    {
                        idList.Add(content);
                    }

                    //if (uint.TryParse(content, out v))
                    //{
                    //    sb.Append("\t[").Append(content).Append("]=");
                    //}
                    //else
                    //{
                    //    PrintDataException(row, firstFieldColumn, "The id must is uint class");
                    //    return;
                    //}

                    AppendLine(sb, range, row, fieldInfos);
                }

                sb.Append("}\n");

                //------------------------------------------------------------------

                if (hasDefault)
                {
                    sb.Append("\nlocal defaults = {");

                    TableFieldInfo fieldInfo;
                    bool first = true;

                    for (int i = 0; i < fieldInfos.Count; i++)
                    {
                        fieldInfo = fieldInfos[i];
                        if (fieldInfo.hasDefault)
                        {
                            if (first)
                                first = false;
                            else
                                sb.Append(", ");

                            sb.Append(GetFieldName(fieldInfo.name)).Append("=").Append(GetFieldData(fieldInfo.defaultV, fieldInfo));
                        }
                    }

                    sb.Append("}\n");

                    sb.Append("defaults.__index = defaults\n\nfor _, v in pairs(").Append(sheetName).Append(") do\n\tsetmetatable(v, defaults)\nend\n");
                }

                sb.Append("\nreturn ").Append(sheetName);

                File.WriteAllText(outputFile, sb.ToString());
            }

            void AppendLine(StringBuilder sb, CellRange range, int row, List<TableFieldInfo> fieldInfos)
            {
                sb.Append("\t{");

                TableFieldInfo fieldinfo;
                bool firstField = true;
                string content;
                for (int i = 0; i < fieldInfos.Count; i++)
                {
                    fieldinfo = fieldInfos[i];
                    content = range[row, fieldinfo.column].DisplayedText;
                    if (!string.IsNullOrEmpty(fieldinfo.name) && (!fieldinfo.hasDefault || fieldinfo.defaultV != content))
                    {
                        if (firstField)
                            firstField = false;
                        else
                            sb.Append(", ");

                        sb.Append(GetFieldName(fieldinfo.name)).Append("=").Append(GetFieldData(content, fieldinfo));
                    }
                }

                if (firstField)
                {
                    sb.Remove(sb.Length - 2, 2);
                    return;
                }

                sb.Append("},\n");
            }

            string GetFieldData(string content, TableFieldInfo fieldInfo)
            {
                switch (fieldInfo.type)
                {
                    case TableFieldType.num:
                        return GetNum(content);
                    case TableFieldType.list:
                        return GetList(content);
                    case TableFieldType.map:
                        return GetMap(content);
                    case TableFieldType.listlist:
                        return GetListList(content);
                    case TableFieldType.sundry:
                        if (content.Contains("&"))
                            return GetMap(content);
                        else if (content.Contains("|"))
                            return GetListList(content);
                        else if (content.Contains("#"))
                            return GetList(content);
                        else
                            return GetNum(content);
                }

                tempSb.Clear();
                tempSb.Append("\"").Append(content).Append("\"");
                return tempSb.ToString();
            }

            static string GetNum(string content)
            {
                float v;
                if (float.TryParse(content, out v))
                {
                    return v.ToString();
                }

                return "0";
            }

            string GetList(string content)
            {
                tempSb.Clear();
                tempSb.Append("{");

                if (!string.IsNullOrEmpty(content))
                {
                    string[] strs = content.Split('#');

                    float v;
                    for (int i = 0; i < strs.Length; i++)
                    {
                        if (string.IsNullOrEmpty(strs[i]))
                        {
                            PrintError(fileName + " List data format is error: " + content);
                            continue;
                        }

                        if (i != 0)
                            tempSb.Append(",");

                        if (float.TryParse(strs[i], out v))
                        {
                            tempSb.Append(strs[i]);
                        }
                        else if (!string.IsNullOrEmpty(strs[i]))
                        {
                            tempSb.Append("\"").Append(strs[i]).Append("\"");
                        }
                    }
                }

                tempSb.Append("}");

                return tempSb.ToString();
            }

            string GetMap(string content)
            {
                tempSb.Clear();
                tempSb.Append("{");

                if (!string.IsNullOrEmpty(content))
                {
                    string[] strs = content.Split('&');
                    string[] pairs;
                    uint v;

                    for (int i = 0; i < strs.Length; i++)
                    {
                        pairs = strs[i].Split('#');
                        if (pairs.Length == 2 && !string.IsNullOrEmpty(pairs[0]) && !string.IsNullOrEmpty(pairs[1]))
                        {
                            if (i != 0)
                                tempSb.Append(",");

                            if (uint.TryParse(pairs[0], out v))
                            {
                                tempSb.Append("[").Append(pairs[0]).Append("]=");
                            }
                            else
                            {
                                tempSb.Append(pairs[0]).Append("=");
                            }

                            if (uint.TryParse(pairs[1], out v))
                            {
                                tempSb.Append(pairs[1]);
                            }
                            else
                            {
                                tempSb.Append("\"").Append(pairs[1]).Append("\"");
                            }
                        }
                        else
                        {
                            PrintError(fileName + " Map data format is error: " + content);
                        }
                    }
                }

                tempSb.Append("}");

                return tempSb.ToString();
            }

            string GetListList(string content)
            {
                tempSb.Clear();
                tempSb.Append("{");

                if (!string.IsNullOrEmpty(content))
                {
                    string[] strs = content.Split('|');
                    string[] strs2;
                    float v;

                    for (int i = 0; i < strs.Length; i++)
                    {
                        if (string.IsNullOrEmpty(strs[i]))
                        {
                            PrintError(fileName + " ListList data format is error: " + content);
                            continue;
                        }

                        if (i != 0)
                            tempSb.Append(",");

                        strs2 = strs[i].Split('#');

                        tempSb.Append("{");

                        for (int k = 0; k < strs2.Length; k++)
                        {
                            if (string.IsNullOrEmpty(strs2[k]))
                            {
                                PrintError(fileName + " ListList data format is error: " + content);
                                continue;
                            }

                            if (k != 0)
                                tempSb.Append(",");

                            if (float.TryParse(strs2[k], out v))
                            {
                                tempSb.Append(strs2[k]);
                            }
                            else
                            {
                                tempSb.Append("\"").Append(strs2[k]).Append("\"");
                            }
                        }

                        tempSb.Append("}");
                    }
                }

                tempSb.Append("}");

                return tempSb.ToString();
            }

            static readonly List<string> LUA_KEYWORDS = new List<string>
            {
                "and", "break", "do", "else", "elseif",
                "end", "false", "for", "function", "if",
                "in", "local", "nil", "not", "or",
                "repeat", "return", "then", "true", "until",
                "while", "goto",
            };

            static bool IsLuaKeyword(string name)
            {
                return LUA_KEYWORDS.Contains(name);
            }

            static string GetFieldName(string name)
            {
                if (IsLuaKeyword(name))
                    return string.Format("[\"{0}\"]", name);
                else
                    return name;
            }
        }

        #region PrintCell
        static void PrintCellRange(CellRange range)
        {
            PrintDividingLine();

            for (int i = 0; i < range.Rows.Length; i++)
            {
                for (int j = 0; j < range.Columns.Length; j++)
                {
                    PrintContent(range[i + 1, j + 1].DisplayedText);
                    PrintDivider();
                }
                Console.Write("\n");
                PrintDividingLine();
            }
        }

        static void PrintContent(string content)
        {
            if (string.IsNullOrEmpty(content))
                Console.BackgroundColor = ConsoleColor.Gray;
            Console.Write(content.PadRight(10));
            Console.BackgroundColor = ConsoleColor.White;
        }

        static void PrintDividingLine()
        {
            PrintGreenInfo("---------------------------------------------------");
        }

        static void PrintDivider()
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.Write(" | ");
            Console.ForegroundColor = ConsoleColor.Black;
        }
        #endregion

        static void PrintGreenInfo(string content)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(content);
            Console.ForegroundColor = ConsoleColor.Black;
        }

        static void End()
        {
            PrintGreenInfo("\n---------------------结束---------------------");
            PrintGreenInfo(string.Format("用时 {0} 秒", (DateTime.Now.Ticks - timeKeeping) / 10000000));

            while (Console.ReadKey().Key != ConsoleKey.Escape) { }
        }

        static void PrintError(string error)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("[error]" + error);
            Console.ForegroundColor = ConsoleColor.Black;
        }

        static void PrintDataException(int row, int column, string error)
        {
            string columnStr = "";
            int val = column - 1;

            while (val >= 0)
            {
                columnStr = (char)(val % 26 + 65) + columnStr;
                val = val / 26 - 1;
            }

            PrintError(string.Format("R:{0} C:{1}({2}) {3}", row, column, columnStr, error));
        }

        static bool IsFileInUse(string fileName, out string error)
        {
            bool inUse = true;
            error = "";
            FileStream fs = null;

            try
            {

                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read,

                FileShare.None);

                inUse = false;
            }
            catch (Exception e)
            {
                error = e.Message;
            }
            finally
            {
                if (fs != null)

                    fs.Close();
            }
            return inUse;//true表示正在使用,false没有使用
        }
    }

    class TableFieldInfo
    {
        public string name;
        public TableFieldType type;
        public int column;
        public string defaultV;
        public bool hasDefault;
    }

    enum TableFieldType
    {
        none,
        num,
        str,
        list,
        map,
        listlist,
        sundry,
    }
}
