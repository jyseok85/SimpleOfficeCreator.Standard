using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace SimpleOfficeCreator.Standard._Samples
{
    internal class Utils
    {
        private Utils() { }
        //private static 인스턴스 객체
        private static readonly Lazy<Utils> _instance = new Lazy<Utils>(() => new Utils());
        //public static 의 객체반환 함수
        public static Utils Instance { get { return _instance.Value; } }

        public string GetWebImage(string url, string defaultValue = "")
        {
            
            try
            {
                WebRequest req = WebRequest.Create(url);
                req.Timeout = 1000;
                using (WebResponse res = req.GetResponse())
                {
                    using (Stream st = res.GetResponseStream())
                    {
                        byte[] buf;
                        using (MemoryStream ms = new MemoryStream())
                        {
                            st.CopyTo(ms);
                            buf = ms.ToArray();
                        }
                        return Convert.ToBase64String(buf);

                    }
                }
            }
            catch (Exception ex)
            {
                if (defaultValue != string.Empty)
                    return defaultValue;
                else
                    throw new Exception("이미지 다운로드에러",ex);
            }
           
        }
    }
}
