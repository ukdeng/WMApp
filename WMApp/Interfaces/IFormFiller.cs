using System;
using WMApp.Models;
using WMApp.ViewModels;

namespace WMApp.Interfaces
{
	public interface IFormFiller
	{
		Dictionary<string,MemoryStream> FillPdf(List<string> files, FormData formData);
	}
}

