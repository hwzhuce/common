using System;
using System.Text;
using System.Net;
using System.IO;
using System.Globalization;
using System.Collections.Specialized;
using System.Web;

namespace Common
{
    /// <summary>
    /// WebClient 帮助器
    /// </summary>
    public class HttpClientHelper
    {
        /// <summary>
        /// 创建默认的 
        /// </summary>
        /// <param name="url">请求的地址</param>
        /// <returns></returns>
        public static HttpWebRequest CreateHttpWebRequest(string url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.AllowAutoRedirect = true;
            //request.Method = "GET";
            request.KeepAlive = true;
            request.Referer = null;
            request.Timeout = 100000; // 默认 100 秒
            return request;
        }

        /// <summary>
        /// 向 HttpWebRequest 请求流写入 POST 数据
        /// </summary>
        /// <param name="request">HttpWebRequest 对象</param>
        /// <param name="postVariables">提交的数据。比如：UserName=bruce&Password=123456 。</param>
        /// <param name="encoding">数据编码，可为 null。如果为 null ，则默认 UTF8 编码 </param>
        public static void WritePostData(HttpWebRequest request, NameValueCollection postVariables, Encoding encoding)
        {
            string postData = GetHttpPostVars(postVariables);
            WritePostData(request, postData, encoding);
        }

        /// <summary>
        /// 向 HttpWebRequest 请求流写入 POST 数据
        /// </summary>
        /// <param name="request">HttpWebRequest 对象</param>
        /// <param name="postData">提交的数据。比如：UserName=bruce&Password=123456 。注意：如果 IsNullOrEmpty 则不会向 HttpWebRequest 请求流中写入数据</param>
        /// <param name="encoding">数据编码，可为 null。如果为 null ，则默认 UTF8 编码 </param>
        public static void WritePostData(HttpWebRequest request, string postData, Encoding encoding)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            if (string.IsNullOrEmpty(postData))
            {
                return;
            }
            if (encoding == null)
            {
                encoding = Encoding.UTF8;
            }
            if (request.Method != "POST")
            {
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded; charset=" + encoding.WebName;
            }
            using (BinaryWriter bw = new BinaryWriter(request.GetRequestStream()))
            {
                bw.Write(encoding.GetBytes(postData));
                bw.Flush();
            }
        }

        /// <summary>
        /// 从 NameValueCollection 中得到 PostData
        /// </summary>
        /// <param name="variables">集合</param>
        /// <returns>PostData</returns>
        public static string GetHttpPostVars(NameValueCollection variables)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < variables.Count; i++)
            {
                string key = variables.GetKey(i);
                string[] values = variables.GetValues(i);
                if (values != null)
                {
                    foreach (string value in values)
                    {
                        builder.AppendFormat("{0}={1}", HttpUtility.UrlEncode(key), HttpUtility.UrlEncode(value));
                    }
                }
                if (i < variables.Count - 1)
                {
                    builder.Append("&");
                }
            }
            return builder.ToString();
        }

        /// <summary>
        /// 根据 HttpWebRequest 请求对象得到响应的 Html 代码
        /// </summary>
        /// <param name="request">HttpWebRequest 对象</param>
        /// <returns></returns>
        public static string GetResponseText(HttpWebRequest request)
        {
            string result = string.Empty;
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            HttpWebResponse response = null;

            try
            {
                response = (HttpWebResponse)request.GetResponse();
            }
            catch (WebException wex)
            {
                if (wex.Response == null)
                {
                    return wex.Message;
                }
                response = (HttpWebResponse)wex.Response;
            }
            using (response)
            {
                using (Stream strem = response.GetResponseStream())
                {
                    MemoryStream ms = null;
                    Byte[] buffer = new Byte[response.ContentLength];
                    int offset = 0, actuallyRead = 0;
                    do
                    {
                        actuallyRead = strem.Read(buffer, offset, buffer.Length - offset);
                        offset += actuallyRead;
                    }
                    while (actuallyRead > 0);
                    result =Encoding.UTF8.GetString(buffer);
                    return result;
                }
            }
        }

        /// <summary>
        /// 发送 Get 请求，得到 Html 代码
        /// </summary>
        /// <param name="url">请求地址</param>
        /// <param name="cookieContainer">Cookie 容器，可为 null。响应时远程服务器返回的 Cookie 也自动包含其中</param>
        /// <returns>Html 代码</returns>
        public static string SendHttpGet(string url, CookieContainer cookieContainer)
        {
            HttpWebRequest request = CreateHttpWebRequest(url);

            if (cookieContainer != null)
            {
                request.CookieContainer = cookieContainer;
            }
            return GetResponseText(request);
        }

        /// <summary>
        /// 发送 POST 请求，得到 Html 代码
        /// </summary>
        /// <param name="url">请求地址</param>
        /// <param name="postData">提交的数据。</param>
        /// <param name="cookieContainer">Cookie 容器，可为 null。响应时远程服务器返回的 Cookie 也自动包含其中</param>
        /// <returns>Html 代码</returns>
        public static string SendHttpPost(string url, string postData, CookieContainer cookieContainer)
        {
            HttpWebRequest request = CreateHttpWebRequest(url);

            if (cookieContainer != null)
            {
                request.CookieContainer = cookieContainer;
            }
            WritePostData(request, postData, Encoding.UTF8);

            return GetResponseText(request);
        }

        /// <summary>
        /// 发送 POST 请求，得到 Html 代码
        /// </summary>
        /// <param name="url">请求地址</param>
        /// <param name="postVariables">提交的数据。</param>
        /// <param name="cookieContainer">Cookie 容器，可为 null。响应时远程服务器返回的 Cookie 也自动包含其中</param>
        /// <returns>Html 代码</returns>
        public static string SendHttpPost(string url, NameValueCollection postVariables, CookieContainer cookieContainer)
        {
            string postData = GetHttpPostVars(postVariables);
            return SendHttpPost(url, postData, cookieContainer);
        }

        /// <summary>
        /// 根据 HttpWebRequest 请求对象得到响应的 Html 代码
        /// </summary>
        /// <param name="request">HttpWebRequest 对象</param>
        /// <returns></returns>
        public static string GetResponseText(string url)
        {
            string result = string.Empty;
            HttpWebRequest request = CreateHttpWebRequest(url);
            HttpWebResponse response = null;

            response = (HttpWebResponse) request.GetResponse();

            using (response)
            {
                using (Stream strem = response.GetResponseStream())
                {
                    MemoryStream ms = null;
                    Byte[] buffer = new Byte[response.ContentLength];
                    int offset = 0, actuallyRead = 0;
                    do
                    {
                        actuallyRead = strem.Read(buffer, offset, buffer.Length - offset);
                        offset += actuallyRead;
                    }
                    while (actuallyRead > 0);
                    result = Encoding.UTF8.GetString(buffer);
                    return result;
                }
            }
        }

        /// <summary>
        /// 根据 HttpWebRequest 请求对象得到响应
        /// </summary>
        /// <param name="url"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static Stream GetResponseStream(string url)
        {
            bool flag = true;
            HttpWebRequest request = CreateHttpWebRequest(url);
            HttpWebResponse response = null;
            response = (HttpWebResponse)request.GetResponse();
            using (response)
            {
                using (Stream strem = response.GetResponseStream())
                {
                    MemoryStream ms = null;
                    Byte[] buffer = new Byte[response.ContentLength];
                    int offset = 0, actuallyRead = 0;
                    do
                    {
                        actuallyRead = strem.Read(buffer, offset, buffer.Length - offset);
                        offset += actuallyRead;
                    }
                    while (actuallyRead > 0);
                    ms = new MemoryStream(buffer);
                    return ms;
                }
            }
        }
    }
}