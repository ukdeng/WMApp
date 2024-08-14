using WMApp.Models;
using WMApp.ViewModels;

namespace WMApp.Helpers
{
	public static class Mapper
	{
		public static FormData MapModelToFormData(string agentName, string agentCode, FormViewModel vm)
        {

			return new FormData()
			{
				Language = vm.Language,
				AgentCode = agentCode,
				AgentName = agentName,
				FirstName = vm.FirstName,
				LastName = vm.LastName,
				Salutation = vm.Salutation,
				DOB = vm.DOB,
				Address = vm.Address,
				City = vm.City,
				Province = vm.Province,
				PostalCode = vm.PostalCode,
				Email = vm.Email,
				Phone = vm.Phone,
				CivilStatus = vm.CivilStatus,
				Dependents = vm.DependentNumber.ToString(),
				SIN = vm.SIN.ToString(),
				Occupation = vm.Occupation,
				OcpSince = vm.Since,
				Employer = vm.Employer,
				RegisteredInvestmentsValue = vm.ValRegInvst,
				NonRegisteredInvestmentsValue = vm.ValNonRegInvst,
				RealEstateValue = vm.ValRealEstate,
				MortgageValue = vm.ValMortgage,
				AnnualIncome = vm.AnnualIncome,
				AccountType = vm.AccountType,
				SubType = vm.SubType,
				FeeType = vm.FeeType,
				Fee = vm.Fee.ToString(),
				ClientName = vm.ClientName
			};
		}
	}
}

