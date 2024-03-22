using Microsoft.AspNetCore.Mvc;
using WMApp.Interfaces;
using WMApp.ViewModels;
using Microsoft.AspNetCore.Mvc.Rendering;
using WMApp.Helpers;

namespace WMApp.Controllers
{
    public class FormController : Controller
    {
        private readonly ILogger<FormController> _logger;
        private readonly IAccessReader _accessReader;

        public FormController(
            ILogger<FormController> logger,
            IAccessReader accessReader)
        {
            _logger = logger;
            _accessReader = accessReader;
        }

        public IActionResult Index()
        {
            var brokers = _accessReader.GetBrokers();
            var viewModel = new FormViewModel()
            {
                Brokers = brokers
            };

            var countries = Countries.GetAll();
            viewModel.SelectLists = new List<SelectListItem>();
            foreach (var item in countries)
            {
                viewModel.SelectLists.Add(new SelectListItem() { Text = item.Name, Value = item.Id});
            }

            return View(viewModel);
        }
    }
}
