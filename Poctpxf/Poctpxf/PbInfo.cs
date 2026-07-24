using Apttpb;

namespace Poctpxf
{
    public class PbInfo : IPhanbo
    {
        public PbInfo(object ngay_ct1, object ngay_ct2, object tk, string ma_kh, string ma_dvcs)
        {
            this.M_NGAY_CT1 = ngay_ct1;
            this.M_NGAY_CT2 = ngay_ct2;
            this.M_TK = tk;
            this.Ma_Dvcs = ma_dvcs;
            this.Ma_kh = ma_kh;
        }

        public object M_NGAY_CT1 { get; set; }

        public object M_NGAY_CT2 { get; set; }

        public object M_TK { get; set; }

        public string Ma_Dvcs { get; set; }

        public string Ma_kh { get; set; }

        public string TitleView { get; set; }
    }
}
