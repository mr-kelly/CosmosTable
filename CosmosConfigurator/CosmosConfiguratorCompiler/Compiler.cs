using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Excel;

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

    public class CompilerConfig
    {
        public string ExcelPath;
        public string Ext = ".bytes";
        public string[] CommentColumnStartsWith = {"Comment", "#"};
    }

    /// <summary>
    /// Compile Excel to TSV
    /// </summary>
    public class Compiler
    {
        private CompilerConfig _config;

        public Compiler(string excelPath)
            : this(new CompilerConfig()
            {
                ExcelPath = excelPath
            })
        {
        }

        public Compiler(CompilerConfig cfg)
        {
            _config = cfg;
        }

        private bool DoCompiler(FileStream stream)
        {
            IExcelDataReader excelReader = null;
            try
            {
                //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            catch (Exception e1)
            {
                try
                {
                    //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                catch (Exception e2)
                {
                    throw new InvalidExcelException("Cannot read Excel File : " + _config.ExcelPath);
                }
            }

            if (excelReader != null)
            {
                using (excelReader)
                {
                    return DoCompilerExcelReader(excelReader);
                }
            }

            return false;
        }

        private bool DoCompilerExcelReader(IExcelDataReader excelReader)
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

            // Header
            int colIndex = 0;
            foreach (DataColumn column in sheet1.Columns)
            {
                var colNameStr = column.ColumnName;
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
                    }
                }
                colIndex++;
            }
            strBuilder.Append("\n");

            // 寻找注释行，1,或2行
            var hasStatementRow = false;
            var regExCheckStatement = new Regex(@"\[(.*)\]");
            foreach (var cellVal in sheet1.Rows[0].ItemArray)
            {
                if ((cellVal is string))
                {
                    var matches = regExCheckStatement.Matches(cellVal.ToString());
                    if (matches.Count > 0)
                        hasStatementRow = true;    
                }
                
                break;
            }

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

            //5. Data Reader methods
            //while (excelReader.Read())
            //{
            //    //excelReader.GetInt32(0);
            //}

            //6. Free resources (IExcelDataReader is IDisposable)


            var fileName = Path.GetFileNameWithoutExtension(_config.ExcelPath);
            File.WriteAllText(string.Format("{0}{1}", fileName, _config.Ext), strBuilder.ToString());
            return true;
        }
        public bool Run()
        {
            using (FileStream stream = File.Open(_config.ExcelPath, FileMode.Open, FileAccess.Read))
            {
                return DoCompiler(stream);
            }
        }
    }
}
