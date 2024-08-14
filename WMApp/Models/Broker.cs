namespace WMApp.Models
{
    public class Broker
    {
        public Broker(string division, string code, string name)
        {
            Division = division;
            Code = code;
            Name = name;
        }

        public string Division { get; set; }

        public string Code { get; set; }

        public string Name { get; set; }
    }
}
