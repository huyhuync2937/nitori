namespace CrmDmEmailcr
{
    public class Info
    {
        public static string fromAddress { get; set; }

        public static string fromName { get; set; }

        public static string fromPass { get; set; }

        public static string subject { get; set; }

        public static string toAddress { get; set; }

        static Info()
        {
            Info.fromAddress = "";
            Info.fromPass = "";
            Info.toAddress = "";
        }
    }
}
