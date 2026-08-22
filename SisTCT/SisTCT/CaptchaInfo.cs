using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SisTCT
{
    public class CaptchaInfo
    {
        internal static readonly CaptchaInfo captchaInfo_0 = new CaptchaInfo()
        {
            State = "Fail"
        };

        public CaptchaInfo()
        {
        }

        public CaptchaInfo(string err)
        {
            this.State = err;
        }

        public CaptchaInfo(string cap, string key, string val)
        {
            this.Captcha = cap;
            this.Key = key;
            this.Value = val;
            this.State = "OK";
        }

        public string Captcha { get; set; }

        public string Key { get; set; }

        public string Value { get; set; }

        public string State { get; set; }
    }
}
