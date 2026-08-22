using System;
using System.Net;

namespace SisTCT
{
    public class WebClientWithTimeout : WebClient
    {
        private int _timeout = 10000;

        public WebClientWithTimeout()
        {
        }

        public WebClientWithTimeout(int timeout)
        {
            this._timeout = timeout;
        }

        protected override WebRequest GetWebRequest(Uri address)
        {
            WebRequest webRequest = base.GetWebRequest(address);
            webRequest.Timeout = this._timeout;
            return webRequest;
        }
    }
}
