using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using Excel;
using DotLiquid;

namespace CosmosConfigurator
{
    /// <summary>
    /// Invali Excel Exception
    /// </summary>
    public class InvalidExcelException : Exception
    {
        public InvalidExcelException(string msg)
            : base(msg)
        {
        }
    }

    public class CodeGentor
    {
        public string TableFilePath { get; set; }
        public string ClassName { get; set; }
        public List<Hash> Fields { get; set; } // column + type
        public List<Hash> Columns2DefaultValus { get; set; } // column + Default Values
        public string PrimaryKey { get; set; }

        public CodeGentor()
        {
            Fields = new List<Hash>();
            Columns2DefaultValus = new List<Hash>();
        }

    }
    public class CompilerConfig
    {
        public string ExportTabExt = ".bytes";

        public string[] CommentColumnStartsWith = {"Comment", "#"};

        public string ExportCodePath = null;//= "CosmosConfigs.cs";

    }

    /// <summary>
    /// Compile Excel to TSV
    /// </summary>
    public class Compiler
    {
        private CompilerConfig _config;

        public Compiler()
            : this(new CompilerConfig()
            {
            })
        {
        }

        public Compiler(CompilerConfig cfg)
        {
            _config = cfg;
        }

        private string DoCompiler(string path, FileStream stream)
        {
            IExcelDataReader excelReader = null;
            try
            {
                //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            catch (Exception)
            {
                try
                {
                    //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                catch (Exception e2)
                {
                    throw new InvalidExcelException("Cannot read Excel File : " + path + e2.Message);
                }
            }

            if (excelReader != null)
            {
                using (excelReader)
                {
                    return DoCompilerExcelReader(path, excelReader);
                }
            }

            return null;
        }


        private string DoCompilerExcelReader(string path, IExcelDataReader excelReader)
        {
            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            //DataSet result = excelReader.AsDataSet();

            //4. DataSet - Create column names from first row
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();

            if (result.Tables.Count <= 0)
                throw new InvalidExcelException("No Sheet!");

            var sheet1 = result.Tables[0];

            var strBuilder = new StringBuilder();

            var ignoreColumns = new HashSet<int>();
            var ignoreRows = new HashSet<int>();
            
            // 寻找注释行，1,或2行
            var hasStatementRow = false;
            var statementRow = sheet1.Rows[0].ItemArray;
            var regExCheckStatement = new Regex(@"\[(.*)\]");
            foreach (var cellVal in statementRow)
            {
                if ((cellVal is string))
                {
                    var matches = regExCheckStatement.Matches(cellVal.ToString());
                    if (matches.Count > 0)
                    {
                        hasStatementRow = true;
                    }
                }

                break;
            }

            // Header
            int colIndex = 0;
            var codeGentor = new CodeGentor();

            foreach (DataColumn column in sheet1.Columns)
            {
                var colNameStr = column.ColumnName.Trim();
                if (!string.IsNullOrEmpty(colNameStr))
                {
                    var isCommentColumn = false;
                    foreach (var commentStartsWith in _config.CommentColumnStartsWith)
                    {
                        if (colNameStr.StartsWith(commentStartsWith))
                        {
                            isCommentColumn = true;
                            break;
                        }
                    }
                    if (isCommentColumn)
                    {
                        ignoreColumns.Add(colIndex);
                    }
                    else
                    {
                        if (colIndex > 0)
                            strBuilder.Append("\t");
                        strBuilder.Append(colNameStr);

                        string typeName = "string";
                        string defaultVal = "";
                        
                        if (hasStatementRow)
                        {
                            var match = regExCheckStatement.Match(statementRow[colIndex].ToString());
                            var attrs = match.Groups[1].ToString().Split(':');
                            // Type
                            
                            if (attrs.Length > 0)
                            {
                                typeName = attrs[0];
                            }
                            // Default Value
                            if (attrs.Length > 1)
                            {
                                defaultVal = attrs[1];
                            }
                            if (attrs.Length > 2)
                            {
                                if (attrs[2] == "pk")
                                {
                                    codeGentor.PrimaryKey = colNameStr;
                                }
                            }

                        }

                        codeGentor.Fields.Add(Hash.FromAnonymousObject(new
                        {
                            Type = typeName,
                            Name = colNameStr,
                            TypeStr = typeName.Replace("[]", "_array"),
                        }));
                        //codeGentor.Columns2DefaultValus.Add(colNameStr, defaultVal);
                    }
                }
                colIndex++;
            }
            strBuilder.Append("\n");

            // Rows
            var rowIndex = 1;
            foreach (DataRow dRow in sheet1.Rows)
            {
                if (hasStatementRow)
                {
                    // 有声明行，忽略第2行
                    if (rowIndex == 2)
                    {
                        rowIndex++;
                        continue;
                        
                    }
                }
                else
                {
                    // 无声明行，忽略第1行
                    if (rowIndex == 1)
                    {
                        rowIndex++;
                        continue;
                    }
                }

                colIndex = 0;
                foreach (var item in dRow.ItemArray)
                {
                    if (ignoreColumns.Contains(colIndex)) // comment column, ignore
                        continue;

                    if (colIndex > 0)
                        strBuilder.Append("\t");
                    strBuilder.Append(item);
                    colIndex++;
                }
                strBuilder.Append("\n");
                rowIndex++;
            }

            var fileName = Path.GetFileNameWithoutExtension(path);
            File.WriteAllText(string.Format("{0}{1}", fileName, _config.ExportTabExt), strBuilder.ToString());


            // 生成代码
            var template = Template.Parse(File.ReadAllText("./GenCode.tpl"));

            codeGentor.ClassName = string.Join("", (from name in fileName.Split('_') select System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(name)).ToArray());
            codeGentor.TableFilePath = path;
            //var codeSb = new StringBuilder();
            //// aa_bb => AaBb
            //var className = 

            //codeSb.Replace("$TABLE_FILE", path);
            //codeSb.Replace("$CLASS_NAME", className);

            //// fields & parse
            //var fieldsSb = new StringBuilder();
            //var parseSb = new StringBuilder();
            //int columnIndex2 = 0;
            //foreach (var kv in codeGentor.columns2Types) // field name => field type
            //{
            //    fieldsSb.AppendLine(string.Format("\tpublic {0} {1};", kv.Value, kv.Key));
            //    parseSb.AppendLine(string.Format("\t\t{0} = Get_{1}(values[{2}], \"\");", kv.Key, kv.Value.Replace("[]", "_array"), columnIndex2));
            //    columnIndex2++;
            //}
            //codeSb.Replace("$FIELDS", fieldsSb.ToString());
            //codeSb.Replace("$PARSE_FIELDS", parseSb.ToString());

            
            //// PrimaryKey
            //codeSb.Replace("$PRIMARY_KEY", codeGentor.primaryKeyColumn ?? "null");
            return template.Render(Hash.FromAnonymousObject(codeGentor));
        } 

        public bool Compile(string path)
        {
            var exportCodes = new StringBuilder();
            exportCodes.AppendLine("using CosmosConfigurator;");

            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                string code = DoCompiler(path, stream);

                exportCodes.Append(code);
            }

            if (!string.IsNullOrEmpty(_config.ExportCodePath))
                File.WriteAllText(_config.ExportCodePath, exportCodes.ToString());

            return true;
        }
    }
}
