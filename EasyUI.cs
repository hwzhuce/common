using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Common
{
    /// <summary>
    /// EasyUI 行内搜索类 SearchOpt.
    /// </summary>
    public class SearchOpt
    {
        /// <summary>
        /// Gets or sets the field.
        /// </summary>
        /// <value>The field.</value>
        public string field { get; set; }
        /// <summary>
        /// Gets or sets the type.
        /// type : label,text,textarea,checkbox,numberbox,validatebox,datebox,combobox,combotree
        /// </summary>
        /// <value>The type.</value>
        public string type { get; set; }
        /// <summary>
        /// Gets or sets the options.
        /// options:{precision:1}
        /// </summary>
        /// <value>The options.</value>
        public string options { get; set; }
        /// <summary>
        /// Gets or sets the op.
        /// op  :  contains,equal,notequal,beginwith,endwith,less,lessorequal,greater,greaterorequal
        /// </summary>
        /// <value>The op.</value>
        public string op { get; set; }

        /// <summary>
        /// 查询参数值
        /// </summary>
        public string value { get; set; }
    }

    /// <summary>
    /// Class DataGrid.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class DataGrid<T>
    {
        public int total { get; set; }

        public List<T> rows { get; set; }
    }
}
