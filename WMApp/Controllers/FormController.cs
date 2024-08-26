using Microsoft.AspNetCore.Mvc;
using WMApp.Interfaces;
using WMApp.ViewModels;
using WMApp.Helpers;
using System.IO.Compression;

namespace WMApp.Controllers
{
    public class FormController : Controller
    {
        private readonly ILogger<FormController> _logger;
        private readonly IExcelReader _excelReader;
        private readonly IFormFiller _formFiller;

        public FormController(
            ILogger<FormController> logger,
            IExcelReader excelReader,
            IFormFiller formFiller)
        {
            _logger = logger;
            _excelReader = excelReader;
            _formFiller = formFiller;
        }

        public IActionResult Index()
        {
            var brokers = _excelReader.GetBrokers();
            var viewModel = new FormViewModel
            {
                LanguageList = SelectListBuilder.BuildLanguageList(),
                BrokerList = SelectListBuilder.BuildBrokerList(brokers),
                CivilStatusList = SelectListBuilder.BuildCivilStatusList(),
                AccountTypeList = SelectListBuilder.BuildAccountTypeList(),
                SubTypeList = SelectListBuilder.BuildSubTypeList(),
                AccountList = SelectListBuilder.BuildAccountList(),
                SalutationList = SelectListBuilder.BuildSalutationList(),
                ClientNameList = SelectListBuilder.BuildClientNameProvidersList(),
                FeeTypes = SelectListBuilder.BuildFeeTypes()
            };

            return View(viewModel);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Index(FormViewModel model)
        {
            if (!ModelState.IsValid)
            {
                //return RedirectToAction("Index");
            }

            // Find the right forms
            var ss = model.Broker.Split("|");
            var prefix = ss[0];
            var agentName = ss[1];
            var agentCode = ss[2];

            var isNom = model.SubType == "0";
            var isFeeBased = model.FeeType == Constants.FeeBased;

            var formData = Mapper.MapModelToFormData(agentName, agentCode, model);
            var files = FormFinder.BuildTargetFormList(isNom, isFeeBased, prefix, formData.ClientName, model.Accounts);
            var streams = _formFiller.FillPdf(files, formData);
            var zipName = $"{formData.FirstName}_{formData.LastName}_{DateTime.Now:yyyy_MM_dd-HH_mm_ss}.zip";

            using MemoryStream ms = new();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                foreach(var stm in streams)
                {
                    var entry = zip.CreateEntry($"{stm.Key}.pdf");
                    using var fileStream = stm.Value;
                    using (var entryStream = entry.Open())
                    {
                        fileStream.CopyTo(entryStream);
                        entryStream.Dispose();
                    }
                    fileStream.Dispose();
                };
            }

            return File(ms.ToArray(), "application/zip", zipName);
        }
    }
}
