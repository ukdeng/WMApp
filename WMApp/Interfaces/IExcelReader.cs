using WMApp.Models;

namespace WMApp.Interfaces
{
    public interface IExcelReader
    {
        List<Broker> GetBrokers();
    }
}
