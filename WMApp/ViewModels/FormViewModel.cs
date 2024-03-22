using Microsoft.AspNetCore.Mvc.Rendering;

namespace WMApp.ViewModels
{
    public class FormViewModel
    {
        public string Selected { get; set; }

        public Dictionary<string, string> Brokers { get; set; }

        public List<SelectListItem> SelectLists { get; set; }
    }
}
