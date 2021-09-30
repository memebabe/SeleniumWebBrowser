using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumWebBrowser
{
    public class ProxyObject
    {
        public string Host { get; set; }
        public int Port { get; set; }
        public string User { get; set; }
        public string Password { get; set; }

        public override string ToString()
        {
            if (string.IsNullOrEmpty(User))
                return $"{Host}:{Port}";
            else
                return $"{Host}:{Port}:{User}:{Password}";
        }

        public static ProxyObject Parse(string proxyString)
        {
            try
            {
                ProxyObject proxyObject = new ProxyObject();
                var strs = proxyString.Split(':');
                if (strs.Length > 1)
                {
                    proxyObject.Host = strs[0];
                    proxyObject.Port = int.Parse(strs[1]);
                    if (strs.Length > 2)
                        proxyObject.User = strs[2];
                    if (strs.Length > 3)
                        proxyObject.Password = strs[3];

                    return proxyObject;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
