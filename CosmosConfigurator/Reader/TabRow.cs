using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CosmosConfigurator
{

    public class TabRow
    {
        public TabRow()
        {
            
        }
        public virtual void Parse(string[] cellStrs)
        {
        }

        public virtual object PrimaryKey
        {
            get
            {
                return null;
            }
        }

        protected string Get_string(string value, string defaultValue)
        {
            if (string.IsNullOrEmpty(value))
                return defaultValue;
            return value;
        }
        protected int Get_int(string value, string defaultValue)
        {
            return int.Parse(Get_string(value, defaultValue));
        }

        protected uint Get_uint(string value, string defaultValue)
        {
            return uint.Parse(Get_string(value, defaultValue));
        }

        protected string[] Get_string_array(string value, string defaultValue)
        {
            var str = Get_string(value, defaultValue);
            return str.Split('|');
        }
    }

    /// <summary>
    /// Default Tab Row
    /// Store All column Values
    /// </summary>
    public class DefaultTabRow : TabRow
    {
        public string[] Values;

        public override void Parse(string[] cellStrs)
        {
            Values = cellStrs;
        }
    }

}
