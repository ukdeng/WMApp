using System;
using Microsoft.AspNetCore.Mvc.Rendering;
using WMApp.Models;

namespace WMApp.Helpers
{
	public static class FormFinder
	{
        public static List<string> BuildTargetFormList(bool isNom, bool feebased, string domain, List<string> accounts)
        {
            List<string> files = new();
            Dictionary<string, string> scopedfiles;

            if (isNom)
            {
                // Nominee
                scopedfiles = NomineeFileList(domain);
            } else
            {
                // Client List
                scopedfiles = ClientNameFileList(domain);
            }


            if(domain == Constants.SPP) {
                files.Add($"PEAK {domain}/{domain}CCO.pdf");

                if (feebased)
                {
                    files.Add($"PEAK {domain}/{domain.ToUpper()}P00.pdf");
                }
            }

            foreach (var acc in accounts)
            {
                if (scopedfiles.ContainsKey(acc))
                {
                    files.Add(scopedfiles[acc]);
                }
            }

            if (domain == Constants.VMP)
            {
                files.Add($"{domain} FORM CODE/{domain}CCO.pdf");

                if (feebased)
                {
                    files.Add($"{domain} FORM CODE/{domain.ToUpper()}P00.pdf");
                }
            }

            return files;
        }

        public static Dictionary<string, string> NomineeFileList(string domain)
        {
            Dictionary<string, string> nomList = new()
            {
                { "rrsp", $"{domain.ToUpper()}R01" },
                { "spousal_rrsp", $"{domain.ToUpper()}R02" },
                { "tfsa", $"{domain.ToUpper()}CE0" },
                { "rrif", $"{domain.ToUpper()}F01" },
                { "spousal_rrif", $"{domain.ToUpper()}F02" },
                { "lira", $"{domain.ToUpper()}CR0" },
                { "locked-in_rrsp", $"{domain.ToUpper()}RI0" },
                { "lif", $"{domain.ToUpper()}FR0" },
                { "prescribed_rrif", $"{domain.ToUpper()}FP0" },
                { "cash_cad", $"{domain.ToUpper()}OC0" },
                { "cash_usd", $"{domain.ToUpper()}OU0" },
                { "spousal_cash", $"{domain.ToUpper()}OC1" },
                { "spousal_cash_usd", $"{domain.ToUpper()}OU1" }
            };

            if (domain == Constants.VMP)
            {
                nomList.Add("marginal_cad", $"{domain.ToUpper()}MC0");
                nomList.Add("marginal_usd", $"{domain.ToUpper()}MU0");
            }

            return nomList;
        }

        public static Dictionary<string, string> ClientNameFileList(string client)
        {
            Dictionary<string, string> cnList = new()
            {
                { "rrsp", $"{client.ToUpper()}R01" },
                { "spousal_rrsp", $"{client.ToUpper()}R02" },
                { "tfsa", $"{client.ToUpper()}CE0" },
                { "rrif", $"{client.ToUpper()}F01" },
                { "spousal_rrif", $"{client.ToUpper()}F02" },
                { "lira", $"{client.ToUpper()}CR0" },
                { "locked-in_rrsp", $"{client.ToUpper()}RI0" },
                { "lif", $"{client.ToUpper()}FR0" },
                { "prescribed_rrif", $"{client.ToUpper()}FP0" },
                { "cash_cad", $"{client.ToUpper()}OC0" },
                { "cash_usd", $"{client.ToUpper()}OU0" },
                { "spousal_cash", $"{client.ToUpper()}OC1" },
                { "spousal_cash_usd", $"{client.ToUpper()}OU1" },
                { "fhsa", $"{client.ToUpper()}CAP" }, // Client name only
                { "resp_indy", $"{client.ToUpper()}REI" }, // Client name only
                { "resp_fam", $"{client.ToUpper()}REF" }, // Client name only
            };

            return cnList;
        }
    }
}

