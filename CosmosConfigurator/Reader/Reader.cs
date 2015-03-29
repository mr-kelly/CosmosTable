using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CosmosConfigurator
{
    public class ReaderConfig
    {
        public string Content;
        public char[] Separators = new char[] { '\t' };
        public Action<string> OnExceptionEvent;
    }

    public class Reader : IEnumerable<Reader.RowInterator>, IDisposable
    {
        private readonly RowInterator _rowInteratorCache;

        private ReaderConfig _config;

        public Reader(string content)
            : this(new ReaderConfig()
                {
                    Content = content
                })
        {
        }

        public Reader(ReaderConfig config)
        {
            _config = config;

            _rowInteratorCache = new RowInterator(this);  // 用來迭代的
            ParseString(_config.Content);
        }


        private int ColCount;  // 列数

        /// <summary>
        /// 表头信息
        /// </summary>
        public class HeaderInfo
        {
            public int ColumnIndex;
            public string HeaderName;
            public string HeaderDef;
        }

        protected Dictionary<string, HeaderInfo> ColIndex = new Dictionary<string, HeaderInfo>();
        protected Dictionary<int, string[]> TabInfo = new Dictionary<int, string[]>();

        public Dictionary<string, HeaderInfo>.KeyCollection HeaderNames
        {
            get { return ColIndex.Keys; }
        }

        // 直接从字符串分析
        public static Reader LoadFromString(string content)
        {
            Reader tabFile = new Reader(content);
            tabFile.ParseString(content);

            return tabFile;
        }

        // 直接从文件, 传入完整目录，跟通过资源管理器自动生成完整目录不一样，给art库用的
        public static Reader LoadFromFile(string fileFullPath)
        {
            using (FileStream fileStream = new FileStream(fileFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            // 不会锁死, 允许其它程序打开
            {

                StreamReader oReader = new StreamReader(fileStream, System.Text.Encoding.UTF8);
                return new Reader(oReader.ReadToEnd());
            }
        }

        protected bool ParseReader(TextReader oReader)
        {
            // 首行
            var headLine = oReader.ReadLine();
            if (headLine == null)
            {
                OnExeption("Head Line null");
                return false;
            }

            var defLine = oReader.ReadLine(); // 声明行
            if (defLine == null)
            {
                OnExeption("Statemen Line (Line2) Null");
                return false;
            }

            var defLineArr = defLine.Split(_config.Separators, StringSplitOptions.None);

            string[] firstLineSplitString = headLine.Split(_config.Separators, StringSplitOptions.None);  // don't remove RemoveEmptyEntries!
            string[] firstLineDef = new string[firstLineSplitString.Length];
            Array.Copy(defLineArr, 0, firstLineDef, 0, defLineArr.Length);  // 拷贝，确保不会超出表头的

            for (int i = 1; i <= firstLineSplitString.Length; i++)
            {
                var headerString = firstLineSplitString[i - 1];

                var headerInfo = new HeaderInfo
                {
                    ColumnIndex = i,
                    HeaderName = headerString,
                    HeaderDef = firstLineDef[i - 1],
                };

                ColIndex[headerInfo.HeaderName] = headerInfo;
            }
            ColCount = firstLineSplitString.Length;  // 標題

            // 读取内容
            string sLine = "";
            int rowIndex = 1; // 从第1行开始
            while (sLine != null)
            {
                sLine = oReader.ReadLine();
                if (sLine != null)
                {
                    string[] splitString1 = sLine.Split(_config.Separators, StringSplitOptions.None);

                    TabInfo[rowIndex] = splitString1;
                    rowIndex++;
                }
            }
            return true;
        }

        protected bool ParseString(string content)
        {
            using (var oReader = new StringReader(content))
            {
                ParseReader(oReader);
            }

            return true;
        }

        // 将当前保存成文件
        public bool Save(string fileName)
        {
            bool result = false;
            StringBuilder sb = new StringBuilder();

            foreach (var header in ColIndex.Values)
                sb.Append(string.Format("{0}\t", header.HeaderName));
            sb.Append("\r\n");

            foreach (var header in ColIndex.Values)
                sb.Append(string.Format("{0}\t", header.HeaderDef));
            sb.Append("\r\n");

            foreach (KeyValuePair<int, string[]> item in TabInfo)
            {
                foreach (string str in item.Value)
                {
                    sb.Append(str);
                    sb.Append('\t');
                }
                sb.Append("\r\n");
            }

            try
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Create))
                {
                    using (StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8))
                    {
                        sw.Write(sb);

                        result = true;
                    }
                }
            }
            catch (IOException e)
            {
                throw new Exception("可能文件正在被Excel打开?" + e.Message);
                result = false;
            }

            return result;
        }

        // 主要的解析函數
        private string _GetString(int row, int column)
        {
            if (column == 0) // 没有此列
                return string.Empty;
            var rowStrings = TabInfo[row];

            return column - 1 >= rowStrings.Length ? "" : rowStrings[column - 1].ToString();
        }

        public string GetString(int row, int column)
        {
            return _GetString(row, column);
        }

        public string GetString(int row, string columnName)
        {
            HeaderInfo headerInfo;
            if (!ColIndex.TryGetValue(columnName, out headerInfo))
                return string.Empty;

            return GetString(row, headerInfo.ColumnIndex);
        }

        public int GetInteger(int row, int column)
        {
            try
            {
                string field = GetString(row, column);
                return (int)float.Parse(field);
            }
            catch
            {
                return 0;
            }
        }

        public int GetInteger(int row, string columnName)
        {
            try
            {
                string field = GetString(row, columnName);
                return (int)float.Parse(field);
            }
            catch
            {
                return 0;
            }
        }

        public uint GetUInteger(int row, int column)
        {
            try
            {
                string field = GetString(row, column);
                return (uint)float.Parse(field);
            }
            catch
            {
                return 0;
            }
        }

        public uint GetUInteger(int row, string columnName)
        {
            try
            {
                string field = GetString(row, columnName);
                return (uint)float.Parse(field);
            }
            catch
            {
                return 0;
            }
        }
        public double GetDouble(int row, int column)
        {
            try
            {
                string field = GetString(row, column);
                return double.Parse(field);
            }
            catch
            {
                return 0;
            }
        }

        public double GetDouble(int row, string columnName)
        {
            try
            {
                string field = GetString(row, columnName);
                return double.Parse(field);
            }
            catch
            {
                return 0;
            }
        }

        public float GetFloat(int row, int column)
        {
            try
            {
                string field = GetString(row, column);
                return float.Parse(field);
            }
            catch
            {
                return 0;
            }
        }

        public float GetFloat(int row, string columnName)
        {
            try
            {
                string field = GetString(row, columnName);
                return float.Parse(field);
            }
            catch
            {
                return 0;
            }
        }

        public bool GetBool(int row, int column)
        {
            int field = GetInteger(row, column);
            return field != 0;
        }

        public bool GetBool(int row, string columnName)
        {
            int field = GetInteger(row, columnName);
            return field != 0;
        }

        public bool HasColumn(string colName)
        {
            return ColIndex.ContainsKey(colName);
        }

        private void OnExeption(string message)
        {
            if (_config.OnExceptionEvent == null)
                throw new Exception(message);
            else
            {
                _config.OnExceptionEvent(message);
            }
        }

        public int NewColumn(string colName, string defineStr = "")
        {
            if (string.IsNullOrEmpty(colName))
                OnExeption("Null Col Name : " + colName);

            var newHeader = new HeaderInfo
            {
                ColumnIndex = ColIndex.Count + 1,
                HeaderName = colName,
                HeaderDef = defineStr,
            };

            ColIndex.Add(colName, newHeader);
            ColCount++;

            return ColCount;
        }

        public int NewRow()
        {
            string[] list = new string[ColCount];
            int rowId = TabInfo.Count + 1;
            TabInfo.Add(rowId, list);
            return rowId;
        }

        public int GetHeight()
        {
            return TabInfo.Count;
        }

        public int GetColumnCount()
        {
            return ColCount;
        }

        public int GetWidth()
        {
            return ColCount;
        }

        public bool SetValue<T>(int row, int column, T value)
        {
            if (row > TabInfo.Count || column > ColCount || row <= 0 || column <= 0)  //  || column > ColIndex.Count
            {
                throw new Exception(string.Format("Wrong row-{0} or column-{1}", row, column));
                return false;
            }
            string content = Convert.ToString(value);
            if (row == 0)
            {
                foreach (var kv in ColIndex)
                {
                    if (kv.Value.ColumnIndex == column)
                    {
                        ColIndex.Remove(kv.Key);
                        ColIndex[content] = kv.Value;
                        break;
                    }
                }
            }
            var rowStrs = TabInfo[row];
            if (column - 1 >= rowStrs.Length) // 超出, 扩充
            {
                var oldRowStrs = rowStrs;
                rowStrs = TabInfo[row] = new string[column];
                oldRowStrs.CopyTo(rowStrs, 0);
            }
            rowStrs[column - 1] = content;
            return true;
        }

        public bool SetValue<T>(int row, string columnName, T value)
        {
            HeaderInfo headerInfo;
            if (!ColIndex.TryGetValue(columnName, out headerInfo))
                return false;

            return SetValue(row, headerInfo.ColumnIndex, value);
        }

        IEnumerator<RowInterator> IEnumerable<RowInterator>.GetEnumerator()
        {
            int rowStart = 1;
            for (int i = rowStart; i <= GetHeight(); i++)
            {
                _rowInteratorCache.Row = i;
                yield return _rowInteratorCache;
            }
        }

        public IEnumerator GetEnumerator()
        {
            int rowStart = 1;
            for (int i = rowStart; i <= GetHeight(); i++)
            {
                _rowInteratorCache.Row = i;
                yield return _rowInteratorCache;
            }
        }

        public class RowInterator  // 一行
        {
            internal Reader TabFile;

            public int Row { get; internal set; }

            internal RowInterator(Reader tabFile)
            {
                TabFile = tabFile;
            }

            public string GetString(string colName)
            {
                return TabFile.GetString(Row, colName);
            }
            public int GetInteger(string colName)
            {
                return TabFile.GetInteger(Row, colName);
            }
        }

        public void Dispose()
        {
            this.ColIndex.Clear();
            this.TabInfo.Clear();
        }

        public void Close()
        {
            Dispose();
        }
    }
}
