using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Excel;

namespace CosmosConfigurator
{
    /// <summary>
    /// Invali Excel Exception
    /// </summary>
    public class InvalidExcelException : Exception
    {
        public InvalidExcelException(string msg) : base(msg)
        {
        }
    }

    public class CompilerConfig
    {
        public string ExcelPath;
        public string Ext = ".bytes";
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
        public bool Run()
        {
            FileStream stream = File.Open(_config.ExcelPath, FileMode.Open, FileAccess.Read);

            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream))
            {
                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                //IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                //DataSet result = excelReader.AsDataSet();

                //4. DataSet - Create column names from first row
                excelReader.IsFirstRowAsColumnNames = true;
                DataSet result = excelReader.AsDataSet();

                if (result.Tables.Count <= 0)
                    throw new InvalidExcelException("No Sheet!");

                var sheet1 = result.Tables[0];

                var strBuilder = new StringBuilder();

                foreach (DataColumn colName in sheet1.Columns)
                {
                    strBuilder.AppendFormat("{0}\t", colName.ColumnName);
                }
                strBuilder.Append("\n");

                foreach (DataRow dRow in sheet1.Rows)
                {
                    foreach (var item in dRow.ItemArray)
                    {
                        strBuilder.AppendFormat("{0}\t", item);
                    }
                    strBuilder.Append("\n");
                }
                //5. Data Reader methods
                //while (excelReader.Read())
                //{
                //    //excelReader.GetInt32(0);
                //}

                //6. Free resources (IExcelDataReader is IDisposable)
                var fileName = Path.GetFileNameWithoutExtension(_config.ExcelPath);
                File.WriteAllText(string.Format("{0}{1}", fileName, _config.Ext), strBuilder.ToString());
                excelReader.Close();
                return true;
            }


            return false;
        }
    }
}
