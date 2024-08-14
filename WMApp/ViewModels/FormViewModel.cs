using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace WMApp.ViewModels
{
    public class FormViewModel
    {
        [Display(Name = Helpers.Constants.Language)]
        public List<SelectListItem>? LanguageList { get; set; }

        [Required]
        public string Language { get; set; }

        [Display(Name = Helpers.Constants.Broker)]
        public List<SelectListItem>? BrokerList { get; set; }

        [Required]
        public string Broker { get; set; }

        [Required] 
        [StringLength(100)]
        [Display(Name = Helpers.Constants.FirstName)]
        public string FirstName { get; set; }

        [Required]
        [StringLength(100)]
        [Display(Name = Helpers.Constants.LastName)]
        public string LastName { get; set; }

        [Required]
        [DataType(DataType.Date)]
        [Display(Name = Helpers.Constants.DOB)]
        public DateTime DOB { get; set; }

        [Required]
        [StringLength(250)]
        [Display(Name = Helpers.Constants.Address)]
        public string Address { get; set; }

        [Required]
        [StringLength(250)]
        [Display(Name = Helpers.Constants.City)]
        public string City { get; set; }

        [Required]
        [StringLength(50)]
        [Display(Name = Helpers.Constants.Province)]
        public string Province { get; set; }

        [Required]
        [StringLength(50)]
        [Display(Name = Helpers.Constants.PostalCode)]
        public string PostalCode { get; set; }

        [Required]
        [StringLength(100)]
        [DataType(DataType.EmailAddress)]
        [Display(Name = Helpers.Constants.Email)]
        public string Email { get; set; }

        [Required]
        [StringLength(50)]
        [Display(Name = Helpers.Constants.Phone)]
        public string Phone { get; set; }

        
        [Display(Name = Helpers.Constants.Civil)]
        public List<SelectListItem>? CivilStatusList { get; set; }

        [Required]
        public string CivilStatus { get; set; }

        [Required]
        [Display(Name = Helpers.Constants.DependentNumber)]
        public int DependentNumber { get; set; }

        [Required]
        [Display(Name = Helpers.Constants.SIN)]
        public int SIN { get; set; }

        [Required]
        [StringLength(100)]
        [Display(Name = Helpers.Constants.Occupation)]
        public string Occupation { get; set; }

        [Required]
        [DataType(DataType.Date)]
        [Display(Name = Helpers.Constants.Since)]
        public DateTime Since { get; set; }

        [Required]
        [StringLength(100)]
        [Display(Name = Helpers.Constants.Employer)]
        public string Employer { get; set; }

        [Required]
        [Range(0, int.MaxValue)]
        [Display(Name = Helpers.Constants.ValRegInvst)]
        public int ValRegInvst { get; set; }

        [Required]
        [Range(0, int.MaxValue)]
        [Display(Name = Helpers.Constants.ValNonRegInvst)]
        public int ValNonRegInvst { get; set; }

        [Required]
        [Range(0, int.MaxValue)]
        [Display(Name = Helpers.Constants.ValRealEstate)]
        public int ValRealEstate { get; set; }

        [Required]
        [Range(0, int.MaxValue)]
        [Display(Name = Helpers.Constants.ValMortgage)]
        public int ValMortgage { get; set; }

        [Required]
        [Range(0, int.MaxValue)]
        [DataType(DataType.Currency)]
        [Display(Name = Helpers.Constants.AnnualIncome)]
        public int AnnualIncome { get; set; }

        // AccountTypeList: Registered|Non-Registered
        [Display(Name = Helpers.Constants.AccountType)]
        public List<SelectListItem>? AccountTypeList { get; set; }

        [Required]
        public string AccountType{ get; set; }

        // SubType: Nominee|Client Name
        [Display(Name = Helpers.Constants.SubType)]
        public List<SelectListItem>? SubTypeList { get; set; }

        [Required]
        public string SubType { get; set; }

        // Actual account, RSP, TFSA etc
        public List<SelectOption> AccountList { get; set; }

        public List<string> Accounts { get; set; }

        // Donmain VMP|SPP
        [Required]
        [Display(Name = Helpers.Constants.Salutation)]
        public List<SelectListItem>? SalutationList { get; set; }

        public string Salutation { get; set; }

        // Not required, front or feebased
        [Display(Name = Helpers.Constants.FeeType)]
        public List<SelectListItem>? FeeTypes { get; set; }
        
        public string FeeType { get; set; }  // FeeBased | Front

        [Display(Name = Helpers.Constants.Fee, Prompt = "%")]
        public int Fee { get; set; }  // x%

        // MFC | CIG |FID | NBC
        [Display(Name = Helpers.Constants.ClientName)]
        public List<SelectListItem>? ClientNameList { get; set; }

        public string ClientName { get; set; }
    }
}
