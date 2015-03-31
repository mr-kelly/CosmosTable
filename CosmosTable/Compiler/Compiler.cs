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

namespace CosmosTable
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

    /// <summary>
    /// 用来进行模板渲染
    /// </summary>
    public class RenderTemplateVars
    {
        public string TabFilePath { get; set; }
        public string ClassName { get; set; }
        public List<RenderFieldVars> FieldsInternal { get; set; } // column + type

        public List<Hash> Fields
        {
            get { return (from f in FieldsInternal select Hash.FromAnonymousObject(f)).ToList(); }
        } 

        public List<Hash> Columns2DefaultValus { get; set; } // column + Default Values
        public string PrimaryKey { get; set; }

        public RenderTemplateVars()
        {
            FieldsInternal = new List<RenderFieldVars>();
            Columns2DefaultValus = new List<Hash>();
        }
    }

    public class RenderFieldVars
    {
        public int Index { get; set; }
        public string Type { get; set; }
        public string Name { get; set; }
        public string DefaultValue { get; set; }
        public string Comment { get; set; }
    }

    public class CompilerConfig
    {
        public string ExportTabExt = ".bytes";

        public string[] CommentColumnStartsWith = {"Comment", "#"};

        public string ExportCodePath = null;//= "CosmosConfigs.cs";

        public string NameSpace = "AppConfigs";
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

        private Hash DoCompiler(string path, FileStream stream, string compileToFilePath = null)
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
                    return DoCompilerExcelReader(path, excelReader, compileToFilePath);
                }
            }

            return null;
        }


        private Hash DoCompilerExcelReader(string path, IExcelDataReader excelReader, string compileToFilePath = null)
        {
            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            //DataSet result = excelReader.AsDataSet();

            //4. DataSet - Create column names from first row
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();

            if (result.Tables.Count <= 0)
                throw new InvalidExcelException("No Sheet!");


            var renderVars = new RenderTemplateVars();
            renderVars.FieldsInternal = new List<RenderFieldVars>();

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
            
            // 获取注释行
            var commentRow = hasStatementRow ? sheet1.Rows[1].ItemArray : sheet1.Rows[0].ItemArray;
            var commentsOfColumns = new List<string>();
            foreach (var cellVal in commentRow)
            {
                commentsOfColumns.Add(cellVal.ToString());
            }

            // Header
            int colIndex = 0;
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
                                    renderVars.PrimaryKey = colNameStr;
                                }
                            }

                        }

                        renderVars.FieldsInternal.Add(new RenderFieldVars
                        {
                            Index = colIndex,
                            Type = typeName,
                            Name = colNameStr,
                            DefaultValue = defaultVal,
                            Comment = commentsOfColumns[colIndex],
                        });
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
            string exportPath;
            if (!string.IsNullOrEmpty(compileToFilePath))
            {
                exportPath = compileToFilePath;
            }
            else
            {
                // use default
                exportPath = string.Format("{0}{1}", fileName, _config.ExportTabExt);
            }
            File.WriteAllText(exportPath, strBuilder.ToString());


            renderVars.ClassName = string.Join("", (from name in fileName.Split('_') select System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(name)).ToArray());
            renderVars.TabFilePath = exportPath;

            return Hash.FromAnonymousObject(renderVars);
        } 

        public bool Compile(string path, string compileToFilePath = null)
        {
            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {

                // 生成代码
                var template = Template.Parse(File.ReadAllText("./GenCode.cs.tpl"));
                var topHash = new Hash();
                topHash["NameSpace"] = _config.NameSpace;
                var files = new List<Hash>();
                topHash["Files"] = files;

                var hash = DoCompiler(path, stream, compileToFilePath);
                files.Add(hash);
                
                if (!string.IsNullOrEmpty(_config.ExportCodePath))
                    File.WriteAllText(_config.ExportCodePath, template.Render(topHash));

            }


            return true;
        }
    }
}
