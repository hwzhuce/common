using System;
using System.Collections.Generic;
using System.Xml;
using Newtonsoft.Json;

namespace Common
{
    public static class WebServiceUtility<T> where T : class
    {
        /// <summary>
        /// 解释接口平台，接口返回的数据
        /// </summary>
        /// <param name="url">接口平台地址</param>
        /// <param name="errorMessage">错误信息</param>
        /// <returns></returns>
        /// <remarks>
        /// T的结构  
        /// class Test
        /// {
        ///   public T dataList{get;set;}
        /// }
        /// </remarks>
        public static List<T> GetData(string url, out string errorMessage)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(url);
            doc.Save(Console.Out);

            XmlElement root = doc.DocumentElement;
            XmlNode xn = root.SelectSingleNode("//return");

            string returnFlag = string.Empty;

            if (xn != null)
            {
                returnFlag = xn.SelectSingleNode("returnFlag").InnerText;
            }

            if (returnFlag.Trim() == "1")
            {
                var errorBody = xn.SelectSingleNode("messageBody").InnerText;
                errorMessage = xn.SelectSingleNode("messageInfor").InnerText + ',' + errorBody.Substring(errorBody.IndexOf(':') + 1);
                return default(List<T>);
            }

            var entityList = new List<T>();

            XmlNodeList xnl = root.GetElementsByTagName("dataList");
            foreach (XmlNode item in xnl)
            {
                var jsonEnity = JsonConvert.SerializeXmlNode(item);
                var entity = JsonConvert.DeserializeObject<T>(jsonEnity);

                entityList.Add(entity);
            }

            errorMessage = string.Empty;

            return entityList;
        }
    }
}
