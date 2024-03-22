using WMApp.Interfaces;

namespace WMApp.Services
{
    public class AccessReader : IAccessReader
    {
        public Dictionary<string, string> GetBrokers () {
            var dic = new Dictionary<string, string>
            {
                { "key", "value" }
            };

            return dic;
        }
    }
}
