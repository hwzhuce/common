using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Data;
using System.IO;
using System.ComponentModel;

namespace Common
{
    public class FTPClient
    {
        private string ftpServerIP = String.Empty;
        private string ftpUser = String.Empty;
        private string ftpPassword = String.Empty;
        private string ftpRootURL = String.Empty;


        public FTPClient(string url, string userid, string password)
        {
            this.ftpServerIP = url;//ftp的IP地址;
            this.ftpUser = userid;//用户名;
            this.ftpPassword = password;//密码;
            this.ftpRootURL = "ftp://" + url + "/";
        }


        public bool fileUpload(Stream locFileStream, string ftpPath, string ftpFileName)
        {
            bool success = false;
            FtpWebRequest ftpWebRequest = null;

            Stream localFileStream = null;
            Stream requestStream = null;

            try
            {
                string uri = ftpRootURL + ftpPath + ftpFileName;

                ftpWebRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(uri));
                ftpWebRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                ftpWebRequest.UseBinary = true;

                ftpWebRequest.KeepAlive = false;
                ftpWebRequest.Method = WebRequestMethods.Ftp.UploadFile;
                ftpWebRequest.ContentLength = locFileStream.Length;

                int buffLength = 2048;
                byte[] buff = new byte[buffLength];
                int contentLen;

                localFileStream = locFileStream;
                requestStream = ftpWebRequest.GetRequestStream();

                contentLen = localFileStream.Read(buff, 0, buffLength);
                while (contentLen != 0)
                {
                    requestStream.Write(buff, 0, contentLen);
                    contentLen = localFileStream.Read(buff, 0, buffLength);
                }

                success = true;
            }
            catch (Exception)
            {
                success = false;
            }
            finally
            {
                if (requestStream != null)
                {
                    requestStream.Close();
                }

                if (localFileStream != null)
                {
                    localFileStream.Close();
                }
            }

            return success;
        }

        /// <summary> 
        /// 上传 
        /// </summary> 
        /// <param name="localFile">本地文件绝对路径</param> 
        /// <param name="ftpPath">上传到ftp的路径</param> 
        /// <param name="ftpFileName">上传到ftp的文件名</param> 
        public bool fileUpload(FileInfo localFile, string ftpPath, string ftpFileName)
        {
            FileStream localFileStream = null;
            localFileStream = localFile.OpenRead();
            return fileUpload(localFileStream, ftpPath, ftpFileName);
            //bool success = false;
            //FtpWebRequest ftpWebRequest = null;

            //FileStream localFileStream = null;
            //Stream requestStream = null;

            //try
            //{
            //    string uri = ftpRootURL + ftpPath + ftpFileName;

            //    ftpWebRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(uri));
            //    ftpWebRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);
            //    ftpWebRequest.UseBinary = true;

            //    ftpWebRequest.KeepAlive = false;
            //    ftpWebRequest.Method = WebRequestMethods.Ftp.UploadFile;
            //    ftpWebRequest.ContentLength = localFile.Length;

            //    int buffLength = 2048;
            //    byte[] buff = new byte[buffLength];
            //    int contentLen;

            //    localFileStream = localFile.OpenRead();
            //    requestStream = ftpWebRequest.GetRequestStream();

            //    contentLen = localFileStream.Read(buff, 0, buffLength);
            //    while (contentLen != 0)
            //    {
            //        requestStream.Write(buff, 0, contentLen);
            //        contentLen = localFileStream.Read(buff, 0, buffLength);
            //    }

            //    success = true;
            //}
            //catch (Exception)
            //{
            //    success = false;
            //}
            //finally
            //{
            //    if (requestStream != null)
            //    {
            //        requestStream.Close();
            //    }

            //    if (localFileStream != null)
            //    {
            //        localFileStream.Close();
            //    }
            //}

            //return success;
        }

        /// <summary> 
        /// 上传文件 
        /// </summary> 
        /// <param name="localPath">本地文件地址(没有文件名)</param> 
        /// <param name="localFileName">本地文件名</param> 
        /// <param name="ftpPath">上传到ftp的路径</param> 
        /// <param name="ftpFileName">上传到ftp的文件名</param> 
        public bool fileUpload(string localPath, string localFileName, string ftpPath, string ftpFileName)
        {
            bool success = false;

            try
            {
                FileInfo localFile = new FileInfo(localPath + localFileName);
                if (localFile.Exists)
                {
                    success = fileUpload(localFile, ftpPath, ftpFileName);
                }
                else
                {
                    success = false;
                }
            }
            catch (Exception)
            {
                success = false;
            }

            return success;
        }

        /// <summary>
        /// FTP 文件下载。（从“sourceFullFilePath”下载到 “targetPhysicsFullFilePath”）
        /// </summary>
        /// <param name="sourceFullFilePath">The source full file path.（FTP文件路径）</param>
        /// <param name="targetPhysicsFullFilePath">The target physics full file path.（下载目标文件路径）</param>
        /// <returns></returns>
        public bool FileDownload(string sourceFullFilePath, string targetPhysicsFullFilePath)
        {
            bool success = false;
            FtpWebRequest ftpWebRequest = null;
            FtpWebResponse ftpWebResponse = null;
            Stream ftpResponseStream = null;
            FileStream outputStream = null;
            try
            {
                //outputStream = new FileStream(targetPhysicsFullFilePath, FileMode.Create);
                ftpWebRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(sourceFullFilePath));
                ftpWebRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                ftpWebRequest.UseBinary = true;
                ftpWebRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                ftpWebResponse = (FtpWebResponse)ftpWebRequest.GetResponse();
                ftpResponseStream = ftpWebResponse.GetResponseStream();
                long contentLength = ftpWebResponse.ContentLength;

                int bufferSize = 2048;
                byte[] buffer = new byte[bufferSize];
                int readCount;
                
                readCount = ftpResponseStream.Read(buffer, 0, bufferSize);
                outputStream = new FileStream(targetPhysicsFullFilePath, FileMode.Create);
                while (readCount > 0)
                {
                    outputStream.Write(buffer, 0, readCount);
                    readCount = ftpResponseStream.Read(buffer, 0, bufferSize);
                }
                success = true;
            }
            catch (Exception e)
            {
                LogWriter.Write(LOG_CATEGORY.WEB_SERVICE, LOG_LEVEL.ERROR, e.Message);
                success = false;
            }
            finally
            {
                if (outputStream != null)
                {
                    outputStream.Close();
                }
                if (ftpResponseStream != null)
                {
                    ftpResponseStream.Close();
                }

                if (ftpWebResponse != null)
                {
                    ftpWebResponse.Close();
                }
            }
            return success;
        }
        /// <summary> 
        /// 下载文件 
        /// </summary> 
        /// <param name="localPath">本地文件地址(没有文件名)</param> 
        /// <param name="localFileName">本地文件名</param> 
        /// <param name="ftpPath">下载的ftp的路径</param> 
        /// <param name="ftpFileName">下载的ftp的文件名</param> 
        public bool fileDownload(string localPath, string localFileName, string ftpPath, string ftpFileName)
        {
            string sourceFullFilePath = ftpRootURL + ftpPath + ftpFileName;
            string targetPhysicsFullFilePath = localPath + localFileName;
            return FileDownload(sourceFullFilePath, targetPhysicsFullFilePath);
        }


        /// <summary> 
        /// 重命名 
        /// </summary> 
        /// <param name="ftpPath">ftp文件路径</param> 
        /// <param name="currentFilename"></param> 
        /// <param name="newFilename"></param> 
        public bool fileRename(string ftpPath, string currentFileName, string newFileName)
        {
            bool success = false;
            FtpWebRequest ftpWebRequest = null;
            FtpWebResponse ftpWebResponse = null;
            Stream ftpResponseStream = null;

            try
            {
                string uri = ftpRootURL + ftpPath + currentFileName;

                ftpWebRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(uri));
                ftpWebRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                ftpWebRequest.UseBinary = true;
                ftpWebRequest.Method = WebRequestMethods.Ftp.Rename;
                ftpWebRequest.RenameTo = newFileName;

                ftpWebResponse = (FtpWebResponse)ftpWebRequest.GetResponse();
                ftpResponseStream = ftpWebResponse.GetResponseStream();
            }
            catch (Exception)
            {
                success = false;
            }
            finally
            {
                if (ftpResponseStream != null)
                {
                    ftpResponseStream.Close();
                }

                if (ftpWebResponse != null)
                {
                    ftpWebResponse.Close();
                }
            }

            return success;
        }


        /// <summary> 
        /// 消除文件 
        /// </summary> 
        /// <param name="filePath"></param> 
        public bool fileDelete(string ftpPath, string ftpName)
        {
            bool success = false;
            FtpWebRequest ftpWebRequest = null;
            FtpWebResponse ftpWebResponse = null;
            Stream ftpResponseStream = null;
            StreamReader streamReader = null;

            try
            {
                string uri = ftpRootURL + ftpPath + ftpName;

                ftpWebRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(uri));
                ftpWebRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                ftpWebRequest.KeepAlive = false;
                ftpWebRequest.Method = WebRequestMethods.Ftp.DeleteFile;

                ftpWebResponse = (FtpWebResponse)ftpWebRequest.GetResponse();
                long size = ftpWebResponse.ContentLength;

                ftpResponseStream = ftpWebResponse.GetResponseStream();
                streamReader = new StreamReader(ftpResponseStream);

                string result = String.Empty;
                result = streamReader.ReadToEnd();



                success = true;
            }
            catch (Exception)
            {
                success = false;
            }
            finally
            {
                if (streamReader != null)
                {
                    streamReader.Close();
                }

                if (ftpResponseStream != null)
                {
                    ftpResponseStream.Close();
                }

                if (ftpWebResponse != null)
                {
                    ftpWebResponse.Close();
                }
            }

            return success;
        }

        /// <summary> 
        /// 文件存在检查 
        /// </summary> 
        /// <param name="ftpPath"></param> 
        /// <param name="ftpName"></param> 
        /// <returns></returns> 
        public bool fileCheckExist(string ftpPath, string ftpName)
        {
            bool success = false;
            FtpWebRequest ftpWebRequest = null;
            WebResponse webResponse = null;
            StreamReader reader = null;

            try
            {
                string url = ftpRootURL + ftpPath;

                ftpWebRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(url));
                ftpWebRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                ftpWebRequest.Method = WebRequestMethods.Ftp.ListDirectory;
                ftpWebRequest.KeepAlive = false;
                webResponse = ftpWebRequest.GetResponse();
                reader = new StreamReader(webResponse.GetResponseStream());
                string line = reader.ReadLine();
                while (line != null)
                {
                    if (line == ftpName)
                    {
                        success = true;
                        break;
                    }
                    line = reader.ReadLine();
                }
            }
            catch (Exception)
            {
                success = false;
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                }

                if (webResponse != null)
                {
                    webResponse.Close();
                }
            }
            return success;
        }
    }
}
