
using Microsoft.CodeAnalysis.CSharp.Syntax;
using WMApp.Models;

namespace WMApp.Helpers
{
    public static class Countries
    {
        public static List<Country> GetAll()
        {
            return new List<Country> {
                new Country() {Id = "AF", Name = "A"},
                new Country() {Id = "BF", Name = "V"},
            };
        }
    }
}
