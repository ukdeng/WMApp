using System;
namespace WMApp.Models
{
	public class FormData
	{
		public string Language { get; set; }

		public string AgentCode { get; set; }

		public string AgentName { get; set; }

		public string FirstName { get; set; }

		public string LastName { get; set; }

		public string Salutation { get; set; }

		public DateTime DOB { get; set; }

		public string Address { get; set; }

		public string City { get; set; }

        public string Province { get; set; }

        public string PostalCode { get; set; }

		public string Email { get; set; }

		public string Phone { get; set; }

		public string CivilStatus { get; set; }

        public string Dependents { get; set; }

		public string SIN { get; set; }

		public string Occupation { get; set; }

		public DateTime OcpSince { get; set; }

        public string Employer { get; set; }

		public int RegisteredInvestmentsValue { get; set; }

		public int NonRegisteredInvestmentsValue { get; set; }

		public int RealEstateValue { get; set; }

		public int MortgageValue { get; set; }

		public int AnnualIncome { get; set; }

        public string AccountType { get; set; }

        public string SubType { get; set; }

        public string FeeType { get; set; }

		public string Fee { get; set; }

		public string ClientName { get; set; }
    }
}

