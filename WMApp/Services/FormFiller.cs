using System;
using WMApp.Helpers;
using WMApp.Interfaces;
using WMApp.Models;
using WMApp.ViewModels;

namespace WMApp.Services
{
	public class FormFiller : IFormFiller
    {
        private const string VmpPath = "wwwroot/files/VMP FORM CODE/";
        private const string SppPath = "wwwroot/files/Peak SPP/";

        private const string DealerCodeVMP = "7682";
        private const string DealerCodeSPP = "7717";

        public Dictionary<string, MemoryStream> FillPdf(List<string> files, FormData formData)
        {
            Dictionary<string, MemoryStream> streams = new();

            foreach (var file in files)
            {
                MemoryStream s;
                switch (file)
                {
                    case "VMPR01":
                        s = FillVMPR01(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPR01":
                        s = FillSPPR01(formData);
                        streams.Add(file, s);

                        break;
                    case "VMPR02":
                        s = FillVMPR02(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPR02":
                        s = FillSPPR02(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPCE0":
                        s = FillVMPCE0(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPCE0":
                        s = FillSPPCE0(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPF01":
                        s = FillVMPF01(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPF01":
                        s = FillSPPF01(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPF02":
                        s = FillVMPF02(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPF02":
                        s = FillSPPF02(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPCR0":
                        s = FillVMPCR0(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPCR0":
                        s = FillSPPCR0(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPRI0":
                        s = FillVMPRI0(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPRI0":
                        s = FillSPPRI0(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPFR0":
                        s = FillVMPFR0(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPFR0":
                        s = FillSPPFR0(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPFP0":
                        s = FillVMPFP0(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPFP0":
                        s = FillSPPFP0(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPOC0":
                        s = FillVMPOC0(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPOC0":
                        s = FillSPPOC0(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPOU0":
                        s = FillVMPOU0(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPOU0":
                        s = FillSPPOU0(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPOC1":
                        s = FillVMPOC1(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPOC1":
                        s = FillSPPOC1(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPOU1":
                        s = FillVMPOU1(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPOU1":
                        s = FillSPPOU1(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPMC0":
                        s = FillVMPMC0(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPMU0":
                        s = FillVMPMU0(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPCC0":
                        s = FillVMPCC0(formData);
                        streams.Add(file, s);
                        break;
                    case "VMPP00":
                        s = FillVMPP00(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPCC0":
                        s = FillSPPCC0(formData);
                        streams.Add(file, s);
                        break;
                    case "SPPP00":
                        s = FillSPPP00(formData);
                        streams.Add(file, s);
                        break;
                    default:
                        break;
                }
            }

            return streams;
        }

        private static MemoryStream FillVMPR01(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPR01.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPR01(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPR01.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPR02(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPR02.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPR02(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPR02.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPCE0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPCE0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("tel cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPCE0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPCE0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("NoRep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("Mr");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mrs");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Ms");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("first_name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("last_name");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("city");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("postal_code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("cellular");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("birth_date");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin_bin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPF01(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPF01.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPF01(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPF01.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPF02(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPF02.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPF02(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPF01.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPCR0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPCR0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPCR0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPCR0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPRI0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPRI0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPRI0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPRI0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPFR0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPFRO.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPFR0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPFR0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPFP0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPFPO.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPFP0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPFP0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPOC0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPOC0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPOC0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPOC0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPOU0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPOU0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPOU0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPOU0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPOC1(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPOC1.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPOC1(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPOC1.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPOU1(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPOU1.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPOU1(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPOU1.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("# rep");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("M");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("Mme");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("Mlle");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("Dr");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPMC0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPMC0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPMU0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPMU0.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("RepCode");
            repCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("salutations#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("salutations#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("salutations#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("salutations#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("prenom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("nom");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("adresse");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("ville");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("code_postal");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cellulaire");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("ddn");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // sin
            var sin = formDocument.Form.FindFormField("sin1");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("rep_name");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPCC0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPCC0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("field_dealer_code");
            dealerCode.Value = DealerCodeVMP;

            // agent code
            var repCode = formDocument.Form.FindFormField("field_representative_code");
            repCode.Value = fd.AgentCode;

            // language
            switch (fd.Language)
            {
                case Constants.French:
                    var lang = formDocument.Form.FindFormField("2_lang#0");
                    lang.Value = "Yes";
                    break;
                case Constants.English:
                    lang = formDocument.Form.FindFormField("2_lang#1");
                    lang.Value = "Yes";
                    break;
                default:
                    break;
            }

            // salutation
            //var salute = formDocument.Form.FindFormField("2_title");
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        salute.Value = "M.";
            //        break;
            //    case Constants.Mrs:
            //        salute.Value = "Mme";
            //        break;
            //    case Constants.Ms:
            //        salute.Value = "Me";
            //        break;
            //    case Constants.Dr:
            //        salute.Value = "Dr";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("2_first_name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("2_last_name");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("2_address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("2_city");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("2_province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("2_postal_code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("2_phone_cell");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("2_dob");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // dependants
            var dependants = formDocument.Form.FindFormField("num_depedants");
            dependants.Value = fd.Dependents;

            // civil
            switch (fd.CivilStatus)
            {
                case Constants.Single:
                    var civlStatus = formDocument.Form.FindFormField("2_civil_status#0");
                    civlStatus.Value = "Yes";
                    break;
                case Constants.Married:
                    civlStatus = formDocument.Form.FindFormField("2_civil_status#1");
                    civlStatus.Value = "Yes";
                    break;
                case Constants.Divorced:
                    civlStatus = formDocument.Form.FindFormField("2_civil_status#2");
                    civlStatus.Value = "Yes";
                    break;
                case Constants.CommonLaw:
                    civlStatus = formDocument.Form.FindFormField("2_civil_status#3");
                    civlStatus.Value = "Yes";
                    break;
                case Constants.Separated:
                    civlStatus = formDocument.Form.FindFormField("2_civil_status#4");
                    civlStatus.Value = "Yes";
                    break;
                case Constants.Widowed:
                    civlStatus = formDocument.Form.FindFormField("2_civil_status#5");
                    civlStatus.Value = "Yes";
                    break;
                default:
                    break;
            }

            // occupation
            var occupation = formDocument.Form.FindFormField("2_occupation");
            occupation.Value = fd.Occupation;

            var since = formDocument.Form.FindFormField("2_years_with_employer");
            since.Value = fd.OcpSince.Year.ToString();

            var employer = formDocument.Form.FindFormField("2_employers_name");
            employer.Value = fd.Employer;

            // sin
            var sin = formDocument.Form.FindFormField("SIN");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("advisor_1_name_1");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillVMPP00(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{VmpPath}VMPP00.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("Code de conseiller");
            repCode.Value = fd.AgentCode;

            // first name
            var firstName = formDocument.Form.FindFormField("Prénom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("Nom");
            lastName.Value = fd.LastName;

            // fee%
            var fee = formDocument.Form.FindFormField("Frais de gestion à percevoir");
            fee.Value = fd.Fee;

            // agent name
            var repName = formDocument.Form.FindFormField("Text9");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPCC0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPCC0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("dealer.cd");
            dealerCode.Value = DealerCodeSPP;

            // agent code
            var repCode = formDocument.Form.FindFormField("advisor.repCd");
            repCode.Value = fd.AgentCode;

            // language
            switch (fd.Language)
            {
                case Constants.French:
                    var lang = formDocument.Form.FindFormField("client.language.cd#0");
                    lang.Value = "Yes";
                    break;
                case Constants.English:
                    lang = formDocument.Form.FindFormField("client.language.cd#1");
                    lang.Value = "Yes";
                    break;
                default:
                    break;
            }

            // salutation
            //var salute = formDocument.Form.FindFormField("client.name.salutation.desc");
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        salute.Value = "M.";
            //        break;
            //    case Constants.Mrs:
            //        salute.Value = "Mme";
            //        break;
            //    case Constants.Ms:
            //        salute.Value = "Me";
            //        break;
            //    case Constants.Dr:
            //        salute.Value = "Dr.";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("client.name.firstName");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("client.name.lastName");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("client.address[1].streetAddress");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("client.address[1].city");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("client.address[1].province.desc");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("client.address[1].postalCode");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("client.phone[23].bracketNumberSlashExt");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("client.birthDate[yyyy/MM/dd]");
            dob.Value = fd.DOB.ToString("yyyy/MM/dd");

            // dependants
            var dependants = formDocument.Form.FindFormField("client.kyc.numberOfDependents");
            dependants.Value = fd.Dependents;

            // civil
            switch (fd.CivilStatus)
            {
                case Constants.Single:
                    var civlStatus = formDocument.Form.FindFormField("client.maritalStatus.cd#0");
                    civlStatus.Value = "Yes";
                    break;
                case Constants.Married:
                    civlStatus = formDocument.Form.FindFormField("client.maritalStatus.cd#1");
                    civlStatus.Value = "Yes";
                    break;
                case Constants.Divorced:
                    civlStatus = formDocument.Form.FindFormField("client.maritalStatus.cd#2");
                    civlStatus.Value = "Yes";
                    break;
                case Constants.CommonLaw:
                    civlStatus = formDocument.Form.FindFormField("client.maritalStatus.cd#3");
                    civlStatus.Value = "Yes";
                    break;
                case Constants.Separated:
                    civlStatus = formDocument.Form.FindFormField("client.maritalStatus.cd#4");
                    civlStatus.Value = "Yes";
                    break;
                case Constants.Widowed:
                    civlStatus = formDocument.Form.FindFormField("client.maritalStatus.cd#5");
                    civlStatus.Value = "Yes";
                    break;
                default:
                    break;
            }

            // occupation
            var occupation = formDocument.Form.FindFormField("client.kyc.career.desc");
            occupation.Value = fd.Occupation;

            var since = formDocument.Form.FindFormField("_.KYCQ4");
            since.Value = fd.OcpSince.Year.ToString();

            var employer = formDocument.Form.FindFormField("client.kyc.employerName");
            employer.Value = fd.Employer;

            // sin
            var sin = formDocument.Form.FindFormField("client.sin.formatted");
            sin.Value = fd.SIN.ToString();

            // agent name
            var repName = formDocument.Form.FindFormField("advisor.name.fullNameWithNoMiddle");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }

        private static MemoryStream FillSPPP00(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{SppPath}SPPP00.pdf");

            // agent code
            var repCode = formDocument.Form.FindFormField("Code de conseiller");
            repCode.Value = fd.AgentCode;

            // first name
            var firstName = formDocument.Form.FindFormField("Prénom");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("Nom");
            lastName.Value = fd.LastName;

            // fee%
            var fee = formDocument.Form.FindFormField("Frais de gestion à percevoir");
            fee.Value = fd.Fee;

            // agent name
            var repName = formDocument.Form.FindFormField("Text10");
            repName.Value = fd.AgentName;

            return formDocument.Stream;
        }
    }
}