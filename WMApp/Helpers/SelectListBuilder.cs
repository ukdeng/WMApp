using Microsoft.AspNetCore.Mvc.Rendering;
using WMApp.Models;
using WMApp.ViewModels;

namespace WMApp.Helpers
{
	public static class SelectListBuilder
	{
        public static List<SelectListItem> BuildLanguageList()
        {
            var listItems = new List<SelectListItem>
            {
                new SelectListItem()
                {
                    Disabled = true,
                    Selected = true,
                    Value = "",
                    Text = Constants.Select
                },
                new SelectListItem()
                {
                    Value = Constants.French,
                    Text = Constants.French
                },
                new SelectListItem()
                {
                    Value = Constants.English,
                    Text = Constants.English
                },
            };

            return listItems;
        }

        public static List<SelectListItem> BuildBrokerList(List<Broker> brokers)
		{
            var listItems = new List<SelectListItem>
            {
                new SelectListItem()
                {
                    Disabled = true,
                    Selected = true,
                    Text = Constants.Select,
                    Value = ""
                }
            };

            foreach (var b in brokers)
			{
				listItems.Add(new SelectListItem()
				{
					Text = $"{b.Division.Trim()}-{b.Name}-{b.Code}",
					Value = $"{b.Division.Trim()}|{b.Name}|{b.Code}"
                });
			}

			return listItems;
        }

        public static List<SelectListItem> BuildCivilStatusList()
        {
            var listItems = new List<SelectListItem>
            {
                new SelectListItem()
                {
                    Disabled = true,
                    Selected = true,
                    Value = "",
                    Text = Constants.Select
                },
                new SelectListItem()
                {
                    Value = Constants.Single,
                    Text = Constants.Single
                },
                new SelectListItem()
                {
                    Value = Constants.Married,
                    Text = Constants.Married
                },
                new SelectListItem()
                {
                    Value = Constants.Divorced,
                    Text = Constants.Divorced
                },
                new SelectListItem()
                {
                    Value = Constants.CommonLaw,
                    Text = Constants.CommonLaw
                },
                new SelectListItem()
                {
                    Value = Constants.Separated,
                    Text = Constants.Separated
                },
                new SelectListItem()
                {
                    Value = Constants.Widowed,
                    Text = Constants.Widowed
                },
            };

            return listItems;
        }

        public static List<SelectListItem> BuildAccountTypeList()
        {
            var listItems = new List<SelectListItem>
            {
                new SelectListItem()
                {
                    Disabled = true,
                    Selected = true,
                    Value = "",
                    Text = Constants.Select
                }
            };

            listItems.AddRange(new List<SelectListItem>()
            {
                new SelectListItem()
                {
                    Text = "Registered",
                    Value = "0"
                },
                new SelectListItem()
                {
                    Text = "Non-Registered",
                    Value = "1"
                },
            });

            return listItems;
        }

        public static List<SelectListItem> BuildSubTypeList()
        {
            var listItems = new List<SelectListItem>
            {
                new SelectListItem()
                {
                    Disabled = true,
                    Selected = true,
                    Value = "",
                    Text = Constants.Select
                }

            };

            listItems.AddRange(new List<SelectListItem>()
            {
                new SelectListItem()
                {
                    Text = "Nominee",
                    Value = "0"
                },
                new SelectListItem()
                {
                    Text = "Client Name",
                    Value = "1"
                },
            });

            return listItems;
        }

        public static List<SelectOption> BuildAccountList()
        {
            return new List<SelectOption>()
                {
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "RRSP",
                        Value = "rrsp",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "SPOUSAL RRSP",
                        Value = "spousal_rrsp",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "TFSA",
                        Value = "tfsa",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "RRIF",
                        Value = "rrif",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "SPOUSAL RRIF",
                        Value = "spousal_rrif",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "LIRA",
                        Value = "lira",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "LOCKED-IN RRSP",
                        Value = "lockedin_rrsp",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "LIF",
                        Value = "lif",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "PRESCRIBED RRIF",
                        Value = "prescribed_rrif",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "FHSA",
                        Value = "fhsa",
                        Class = "cl-only"
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "CASH CAD",
                        Value = "cash_cad",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "CASH USD",
                        Value = "cash_usd",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "SPOUSAL CASH",
                        Value = "spousal_cash",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "SPOUSAL CASH USD",
                        Value = "spousal_cash_usd",
                        Class = ""
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "MARGINAL CAD",
                        Value = "marginal_cad",
                        Class = "nom-only"
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "MARGINAL USD",
                        Value = "marginal_usd",
                        Class = "nom-only"
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "RESP INDIVIDUAL",
                        Value = "resp_individual",
                        Class = "cl-only"
                    },
                    new SelectOption()
                    {
                        IsChecked = false,
                        Description = "RESP FAMILY",
                        Value = "resp_family",
                        Class = "cl-only"
                    },
                };
        }

        public static List<SelectListItem> BuildSalutationList()
        {
            var listItems = new List<SelectListItem>
            {
                new SelectListItem()
                {
                    Disabled = true,
                    Selected = true,
                    Value = "",
                    Text = Constants.Select
                }

            };

            listItems.AddRange(new List<SelectListItem>()
            {
                new SelectListItem()
                {
                    Text = Constants.Mr,
                    Value = Constants.Mr
                },
                new SelectListItem()
                {
                    Text = Constants.Mrs,
                    Value = Constants.Mrs
                },
                new SelectListItem()
                {
                    Text = Constants.Ms,
                    Value = Constants.Ms
                },
                new SelectListItem()
                {
                    Text = Constants.Dr,
                    Value = Constants.Dr
                },
            });

            return listItems;
        }

        public static List<SelectListItem> BuildFeeTypes()
        {
            var listItems = new List<SelectListItem>
            {
                new SelectListItem()
                {
                    Disabled = true,
                    Selected = true,
                    Value = "",
                    Text = Constants.Select
                }

            };

            listItems.AddRange(new List<SelectListItem>()
            {
                new SelectListItem()
                {
                    Text = Constants.FeeBased,
                    Value = Constants.FeeBased
                },
                new SelectListItem()
                {
                    Text = Constants.Front,
                    Value = Constants.Front
                },
            });

            return listItems;
        }

        public static List<SelectListItem> BuildClientNameProvidersList()
        {
            var listItems = new List<SelectListItem>
            {
                new SelectListItem()
                {
                    Disabled = true,
                    Selected = true,
                    Value = "",
                    Text = Constants.Select
                }

            };

            listItems.AddRange(new List<SelectListItem>()
            {
                new SelectListItem()
                {
                    Text = Constants.MFC,
                    Value = Constants.MFC
                },
                new SelectListItem()
                {
                    Text = Constants.CIG,
                    Value = Constants.CIG
                },
                new SelectListItem()
                {
                    Text = Constants.FID,
                    Value = Constants.FID
                },
                new SelectListItem()
                {
                    Text = Constants.NBC,
                    Value = Constants.NBC
                },
            });

            return listItems;
        }
    }
}