
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security;
using System.Runtime.InteropServices;
using NSoup.Safety;

namespace Common
{
    /// <summary>
    /// 处理字符串中特殊字符，进行安全验证
    /// </summary>
   public static class FilterStringValue
    {
       /// <summary>
        /// 替换sql脚本关键字为""
       /// </summary>
       /// <param name="strValue">待过滤替换原字符串</param>
       /// <returns></returns>
       public static string FilterKeyWord(string strValue)
       {
           if (string.IsNullOrEmpty(strValue))
           {
               return string.Empty;
           }
           strValue = strValue.ToLower();
           strValue = strValue.Replace("=", "");
           strValue = strValue.Replace("'", "");
           strValue = strValue.Replace(";", "");
           strValue = strValue.Replace(" or ", "");
           strValue = strValue.Replace("select", "");
           strValue = strValue.Replace("update", "");
           strValue = strValue.Replace("insert", "");
           strValue = strValue.Replace("delete", "");
           strValue = strValue.Replace("declare", "");
           strValue = strValue.Replace("exec", "");
           strValue = strValue.Replace("drop", "");
           strValue = strValue.Replace("create", "");
           strValue = strValue.Replace("%", "");
           strValue = strValue.Replace("--", "");
           strValue = strValue.Replace("create","");
           return strValue;
       }

       /// <summary>
       /// 过滤文件名字或路径中含不合法的字符
       /// </summary>
       /// <param name="fileName">文件名</param>
       /// <returns></returns>
       public static string FilterFileName(string fileName)
       {
           if (string.IsNullOrEmpty(fileName))
           {
               return string.Empty;
           }
          return fileName.Replace("../", "").Replace("/", "").Replace("..\\", "");
       }

       /// <summary>
       /// 采用白名单允许富文本编辑器指定标签保存入库，其余标签进行过滤，只保留文本值
       /// </summary>
       /// <param name="content">待过滤的字符串</param>
       /// <returns></returns>
       public static string ConvertSpecialChar(string content)
       {
           if (string.IsNullOrEmpty(content))
           {
               return string.Empty;
           }
           content = content.Trim();
           //输入文本中含<input><img>等标签时，需将&gt;&lt;等转义后过滤掉
           content = HtmlChartReplace(content);
           Whitelist tagsAttributes = Whitelist.None.AddTags("table", "thead", "th", "tr", "td", "em", "tbody", "strong", "span", "div", "br", "p");
           tagsAttributes.AddAttributes("table", "cellpadding", "cellspacing");
           tagsAttributes.AddAttributes("th", "style");
           tagsAttributes.AddAttributes("tr", "style", "contenteditable");
           tagsAttributes.AddAttributes("td", "style");
           tagsAttributes.AddAttributes("p", "class", "contenteditable");
           tagsAttributes.AddAttributes("div", "style");
           content = NSoup.NSoupClient.Clean(content, tagsAttributes);
           return content;
       }

       /// <summary>
       /// 将部分字符进行转义
       /// </summary>
       /// <param name="content"></param>
       /// <returns></returns>
       private static string HtmlChartReplace(string content)
       {
           if (string.IsNullOrEmpty(content))
           {
               return string.Empty;
           }
           content = content.Replace("/&amp", "&");
           content = content.Replace("&gt;", ">");
           content = content.Replace("&lt;", "<");
           content = content.Replace("&nbsp;", " ");
           content = content.Replace("&#39;", "\'");
           content = content.Replace("&quot;", "\"");
           return content;
       }
    }
}
