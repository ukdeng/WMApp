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
        private const string MfcPath = "wwwroot/files/MFC/";
        private const string CigPath = "wwwroot/files/CI/";
        private const string FidPath = "wwwroot/files/FID/";

        private const string DealerCodeVMP = "7682";
        private const string DealerCodeSPP = "7717";
        private const string DealerCodeCIG = "";
        private const string DealerAccNumCIG = "";
        private const string DealerCodeFID = "7717";

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
                    case "MFCR01":
                        s = FillMFCR01(formData);
                        streams.Add(file, s);
                        break;
                    case "MFCR02":
                        s = FillMFCR02(formData);
                        streams.Add(file, s);
                        break;
                    case "MFCCE0":
                        s = FillMFCCE0(formData);
                        streams.Add(file, s);
                        break;
                    case "MFCF01":
                        s = FillMFCF01(formData);
                        streams.Add(file, s);
                        break;
                    case "MFCF02":
                        s = FillMFCF02(formData);
                        streams.Add(file, s);
                        break;
                    case "MFCCR0":
                        s = FillMFCCR0(formData);
                        streams.Add(file, s);
                        break;
                    case "MFCRI0":
                        s = FillMFCRI0(formData);
                        streams.Add(file, s);
                        break;
                    case "MFCFR0":
                        s = FillMFCFR0(formData);
                        streams.Add(file, s);
                        break;
                    //case "MFCFP0":
                    //    s = FillMFCFP0(formData);
                    //    streams.Add(file, s);
                    //    break;
                    //case "MFCOC0":
                    //    s = FillMFCOC0(formData);
                    //    streams.Add(file, s);
                    //    break;
                    //case "MFCOU0":
                    //    s = FillMFCOU0(formData);
                    //    streams.Add(file, s);
                    //    break;
                    //case "MFCOC1":
                    //    s = FillMFCOC1(formData);
                    //    streams.Add(file, s);
                    //    break;
                    //case "MFCOU1":
                    //    s = FillMFCOU1(formData);
                    //    streams.Add(file, s);
                    //    break;
                    case "MFCCAP":
                        s = FillMFCCAP(formData);
                        streams.Add(file, s);
                        break;
                    case "MFCREI":
                        s = FillMFCREI(formData);
                        streams.Add(file, s);
                        break;
                    case "MFCREF":
                        s = FillMFCREF(formData);
                        streams.Add(file, s);
                        break;

                    case "CIGR01":
                        s = FillCIGR01(formData);
                        streams.Add(file, s);
                        break;
                    case "CIGR02":
                        s = FillCIGR02(formData);
                        streams.Add(file, s);
                        break;
                    case "CIGCE0":
                        s = FillCIGCE0(formData);
                        streams.Add(file, s);
                        break;
                    case "CIGF01":
                        s = FillCIGF01(formData);
                        streams.Add(file, s);
                        break;
                    case "CIGF02":
                        s = FillCIGF02(formData);
                        streams.Add(file, s);
                        break;
                    case "CIGCR0":
                        s = FillCIGCR0(formData);
                        streams.Add(file, s);
                        break;
                    case "CIGRI0":
                        s = FillCIGRI0(formData);
                        streams.Add(file, s);
                        break;
                    case "CIGFR0":
                        s = FillCIGFR0(formData);
                        streams.Add(file, s);
                        break;
                    case "CIGFP0":
                        s = FillCIGFP0(formData);
                        streams.Add(file, s);
                        break;
                    case "CIGOC0":
                        s = FillCIGOC0(formData);
                        streams.Add(file, s);
                        break;
                    ////case "CIGOU0":
                    ////    s = FillCIGOU0(formData);
                    ////    streams.Add(file, s);
                    ////    break;
                    case "CIGOC1":
                        s = FillCIGOC1(formData);
                        streams.Add(file, s);
                        break;
                    ////case "CIGOU1":
                    ////    s = FillCIGOU1(formData);
                    ////    streams.Add(file, s);
                    ////    break;
                    case "CIGCAP":
                        s = FillCIGCAP(formData);
                        streams.Add(file, s);
                        break;
                    case "CIGREI":
                        s = FillCIGREI(formData);
                        streams.Add(file, s);
                        break;
                    case "CIGREF":
                        s = FillCIGREF(formData);
                        streams.Add(file, s);
                        break;


                    case "FIDR01":
                        s = FillFIDR01(formData);
                        streams.Add(file, s);
                        break;
                    case "FIDR02":
                        s = FillFIDR02(formData);
                        streams.Add(file, s);
                        break;
                    case "FIDCE0":
                        s = FillFIDCE0(formData);
                        streams.Add(file, s);
                        break;
                    case "FIDF01":
                        s = FillFIDF01(formData);
                        streams.Add(file, s);
                        break;
                    case "FIDF02":
                        s = FillFIDF02(formData);
                        streams.Add(file, s);
                        break;
                    case "FIDCR0":
                        s = FillFIDCR0(formData);
                        streams.Add(file, s);
                        break;
                    case "FIDRI0":
                        s = FillFIDRI0(formData);
                        streams.Add(file, s);
                        break;
                    case "FIDFR0":
                        s = FillFIDFR0(formData);
                        streams.Add(file, s);
                        break;
                    ////case "FIDFP0":
                    ////    s = FillMFCFP0(formData);
                    ////    streams.Add(file, s);
                    ////    break;
                    case "FIDOC0":
                        s = FillFIDOC0(formData);
                        streams.Add(file, s);
                        break;
                    ////case "FIDOU0":
                    ////    s = FillMFCOU0(formData);
                    ////    streams.Add(file, s);
                    ////    break;
                    ////case "FIDOC1":
                    ////    s = FillMFCOC1(formData);
                    ////    streams.Add(file, s);
                    ////    break;
                    ////case "FIDOU1":
                    ////    s = FillMFCOU1(formData);
                    ////    streams.Add(file, s);
                    ////    break;
                    case "FIDCAP":
                        s = FillFIDCAP(formData);
                        streams.Add(file, s);
                        break;
                    case "FIDREI":
                        s = FillFIDREI(formData);
                        streams.Add(file, s);
                        break;
                    case "FIDREF":
                        s = FillFIDREF(formData);
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

        private static MemoryStream FillMFCR01(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{MfcPath}MFCR01.pdf");

            // salutation
            var salute = formDocument.Form.FindFormField("Text33");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "1";
                    break;
                case Constants.Mrs:
                    salute.Value = "2";
                    break;
                case Constants.Ms:
                    salute.Value = "3";
                    break;
                case Constants.Dr:
                    salute.Value = "5";
                    break;
                default:
                    break;
            }

            // last name
            var lastName = formDocument.Form.FindFormField("Text34");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("Text36");
            firstName.Value = fd.FirstName;

            // address
            var address = formDocument.Form.FindFormField("Text38");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Text43");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Text73");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("Text72");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Text37");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("Text39");
            email.Value = fd.Email;

            // dob
            var dob = formDocument.Form.FindFormField("Date45_af_date");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("Text42");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr  = formDocument.Form.FindFormField("Check Box36");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("Check Box35");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillMFCR02(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{MfcPath}MFCR02.pdf");

            // salutation
            var salute = formDocument.Form.FindFormField("Text33");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "1";
                    break;
                case Constants.Mrs:
                    salute.Value = "2";
                    break;
                case Constants.Ms:
                    salute.Value = "3";
                    break;
                case Constants.Dr:
                    salute.Value = "5";
                    break;
                default:
                    break;
            }

            // last name
            var lastName = formDocument.Form.FindFormField("Text34");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("Text36");
            firstName.Value = fd.FirstName;

            // address
            var address = formDocument.Form.FindFormField("Text38");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Text43");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Text73");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("Text72");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Text35");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("Text39");
            email.Value = fd.Email;

            // dob
            var dob = formDocument.Form.FindFormField("Date45_af_date");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("Text42");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("Check Box36");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("Check Box35");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillMFCCE0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{MfcPath}MFCCE0.pdf");

            // salutation
            var salute = formDocument.Form.FindFormField("2 RENSEIGNEMENTS SUR LE TITULAIRE DU COMPTE En caractères dimprimerie");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "1";
                    break;
                case Constants.Mrs:
                    salute.Value = "2";
                    break;
                case Constants.Ms:
                    salute.Value = "3";
                    break;
                case Constants.Dr:
                    salute.Value = "5";
                    break;
                default:
                    break;
            }

            // last name
            var lastName = formDocument.Form.FindFormField("Nom de famille");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("Prénom");
            firstName.Value = fd.FirstName;

            // address
            var address = formDocument.Form.FindFormField("Adresse_1");
            address.Value = fd.Address;

            // postal code
            var postalCode = formDocument.Form.FindFormField("Adresse_2");
            postalCode.Value = $"                        {fd.PostalCode}";


            // city province
            var cp = formDocument.Form.FindFormField("Ville");
            cp.Value = $"{fd.City}             {fd.Province}";

            // cell
            var cell = formDocument.Form.FindFormField("Téléphone travail");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("Courriel");
            email.Value = fd.Email;

            //sociale requis sin
            var sin = formDocument.Form.FindFormField("Numéro dassurance sociale requis");
            sin.Value = fd.SIN.ToString();

            // dob
            var dd = formDocument.Form.FindFormField("Date de naissance-1");
            dd.Value = fd.DOB.Day.ToString("dd");

            var mm = formDocument.Form.FindFormField("Date de naissance-2");
            mm.Value = fd.DOB.Month.ToString("MM");

            var dob = formDocument.Form.FindFormField("Date de naissance-3");
            dob.Value = fd.DOB.Year.ToString("yyyy");

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("Français");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("Anglais");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillMFCF01(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{MfcPath}MFCF01.pdf");

            // salutation
            var salute = formDocument.Form.FindFormField("Text33");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "1";
                    break;
                case Constants.Mrs:
                    salute.Value = "2";
                    break;
                case Constants.Ms:
                    salute.Value = "3";
                    break;
                case Constants.Dr:
                    salute.Value = "5";
                    break;
                default:
                    break;
            }

            // last name
            var lastName = formDocument.Form.FindFormField("Text34");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("Text36");
            firstName.Value = fd.FirstName;

            // address
            var address = formDocument.Form.FindFormField("Text38");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Text43");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Text73");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("Text72");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Text37");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("Text39");
            email.Value = fd.Email;

            // dob
            var dob = formDocument.Form.FindFormField("Date45_af_date");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("Text42");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("Check Box36");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("Check Box35");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillMFCF02(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{MfcPath}MFCF02.pdf");

            // salutation
            var salute = formDocument.Form.FindFormField("Text33");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "1";
                    break;
                case Constants.Mrs:
                    salute.Value = "2";
                    break;
                case Constants.Ms:
                    salute.Value = "3";
                    break;
                case Constants.Dr:
                    salute.Value = "5";
                    break;
                default:
                    break;
            }

            // last name
            var lastName = formDocument.Form.FindFormField("Text34");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("Text36");
            firstName.Value = fd.FirstName;

            // address
            var address = formDocument.Form.FindFormField("Text38");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Text43");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Text73");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("Text72");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Text37");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("Text39");
            email.Value = fd.Email;

            // dob
            var dob = formDocument.Form.FindFormField("Date45_af_date");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("Text42");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("Check Box36");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("Check Box35");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillMFCCR0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{MfcPath}MFCCR0.pdf");

            // salutation
            var salute = formDocument.Form.FindFormField("Text33");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "1";
                    break;
                case Constants.Mrs:
                    salute.Value = "2";
                    break;
                case Constants.Ms:
                    salute.Value = "3";
                    break;
                case Constants.Dr:
                    salute.Value = "5";
                    break;
                default:
                    break;
            }

            // last name
            var lastName = formDocument.Form.FindFormField("Text34");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("Text36");
            firstName.Value = fd.FirstName;

            // address
            var address = formDocument.Form.FindFormField("Text38");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Text43");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Text73");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("Text72");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Text37");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("Text39");
            email.Value = fd.Email;

            // dob
            var dob = formDocument.Form.FindFormField("Date45_af_date");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("Text42");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("Check Box36");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("Check Box35");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillMFCRI0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{MfcPath}MFCRI0.pdf");

            // salutation
            var salute = formDocument.Form.FindFormField("Text33");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "1";
                    break;
                case Constants.Mrs:
                    salute.Value = "2";
                    break;
                case Constants.Ms:
                    salute.Value = "3";
                    break;
                case Constants.Dr:
                    salute.Value = "5";
                    break;
                default:
                    break;
            }

            // last name
            var lastName = formDocument.Form.FindFormField("Text34");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("Text36");
            firstName.Value = fd.FirstName;

            // address
            var address = formDocument.Form.FindFormField("Text38");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Text43");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Text73");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("Text72");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Text37");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("Text39");
            email.Value = fd.Email;

            // dob
            var dob = formDocument.Form.FindFormField("Date45_af_date");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("Text42");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("Check Box36");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("Check Box35");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillMFCFR0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{MfcPath}MFCFR0.pdf");

            // salutation
            var salute = formDocument.Form.FindFormField("Text33");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "1";
                    break;
                case Constants.Mrs:
                    salute.Value = "2";
                    break;
                case Constants.Ms:
                    salute.Value = "3";
                    break;
                case Constants.Dr:
                    salute.Value = "5";
                    break;
                default:
                    break;
            }

            // last name
            var lastName = formDocument.Form.FindFormField("Text34");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("Text36");
            firstName.Value = fd.FirstName;

            // address
            var address = formDocument.Form.FindFormField("Text38");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Text43");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Text73");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("Text72");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Text37");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("Text39");
            email.Value = fd.Email;

            // dob
            var dob = formDocument.Form.FindFormField("Date45_af_date");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("Text42");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("Check Box36");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("Check Box35");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillMFCCAP(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{MfcPath}MFCCAP.pdf");

            // salutation
            var salute = formDocument.Form.FindFormField("Text Field 3");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "1";
                    break;
                case Constants.Mrs:
                    salute.Value = "2";
                    break;
                case Constants.Ms:
                    salute.Value = "3";
                    break;
                case Constants.Dr:
                    salute.Value = "5";
                    break;
                default:
                    break;
            }

            // last name
            var lastName = formDocument.Form.FindFormField("Text Field 4");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("Text Field 5");
            firstName.Value = fd.FirstName;

            // address
            var address = formDocument.Form.FindFormField("Text Field 6");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Text Field 10");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Text Field 11");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("Text Field 9");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Text Field 13");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("Text Field 14");
            email.Value = fd.Email;

            // dob
            var dob = formDocument.Form.FindFormField("Text Field 16");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("Text Field 15");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("Check Box 5");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("Check Box 4");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillMFCREI(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{MfcPath}MFCREI.pdf");

            // salutation
            var salute = formDocument.Form.FindFormField("Text Field 2");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "1";
                    break;
                case Constants.Mrs:
                    salute.Value = "2";
                    break;
                case Constants.Ms:
                    salute.Value = "3";
                    break;
                case Constants.Dr:
                    salute.Value = "5";
                    break;
                default:
                    break;
            }

            // last name
            var lastName = formDocument.Form.FindFormField("Text Field 3");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("Text Field 4");
            firstName.Value = fd.FirstName;

            // address postal
            var ap = formDocument.Form.FindFormField("Text Field 5");
            ap.Value = $"{fd.Address}        {fd.PostalCode}";

            // city province
            var cp = formDocument.Form.FindFormField("Text Field 6");
            cp.Value = $"{fd.City}           {fd.Province}";

            // cell
            var cell = formDocument.Form.FindFormField("Text Field 13");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("Text Field 14");
            email.Value = fd.Email;

            // dob
            var dd = formDocument.Form.FindFormField("Text Field 18a");
            dd.Value = fd.DOB.Day.ToString("dd");

            var mm = formDocument.Form.FindFormField("Text Field 18b");
            mm.Value = fd.DOB.Month.ToString("MM");

            var dob = formDocument.Form.FindFormField("Text Field 18c");
            dob.Value = fd.DOB.Year.ToString("yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("Text Field 19");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("Français");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("Anglais");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillMFCREF(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{MfcPath}MFCREF.pdf");

            // salutation
            var salute = formDocument.Form.FindFormField("IDENTIFICATION DU SOUSCRIPTEUR  OBLIGATOIRE En lettres moulées SVP");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "1";
                    break;
                case Constants.Mrs:
                    salute.Value = "2";
                    break;
                case Constants.Ms:
                    salute.Value = "3";
                    break;
                case Constants.Dr:
                    salute.Value = "5";
                    break;
                default:
                    break;
            }

            // last name
            var lastName = formDocument.Form.FindFormField("Nom de famille");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("Prénom et initiale");
            firstName.Value = fd.FirstName;

            // address postal
            var ap = formDocument.Form.FindFormField("Adresse");
            ap.Value = $"{fd.Address}        {fd.PostalCode}";

            // city province
            var cp = formDocument.Form.FindFormField("Ville");
            cp.Value = $"{fd.City}           {fd.Province}";

            // cell
            var cell = formDocument.Form.FindFormField("Téléphone travail");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("Courriel");
            email.Value = fd.Email;

            // dob
            var dd = formDocument.Form.FindFormField("Birth Date-1");
            dd.Value = fd.DOB.Day.ToString("dd");

            var mm = formDocument.Form.FindFormField("Birth Date-2");
            mm.Value = fd.DOB.Month.ToString("MM");

            var dob = formDocument.Form.FindFormField("Birth Date-3");
            dob.Value = fd.DOB.Year.ToString("yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("NAS");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("Français");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("Anglais");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGR01(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CIR01.pdf");

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // last name
            var lastName = formDocument.Form.FindFormField("03 Unitholder Last Name");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("03 Unitholder Info First Name");
            firstName.Value = fd.FirstName;

            // address postal
            var address = formDocument.Form.FindFormField("03 Unitholder Street Addr");
            address.Value = fd.Address;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Unitholder Postal Code");
            postalCode.Value = fd.PostalCode;

            // city
            var city = formDocument.Form.FindFormField("03 Unitholder City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Unitholder Province");
            province.Value = fd.Province;

            // cell
            var cell = formDocument.Form.FindFormField("03 Unitholder Mobile No");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("03 Unitholder Email Address");
            email.Value = fd.Email;

            // dob
            var dob = formDocument.Form.FindFormField("03 Unitholder DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("03 Unitholder SIN Bus No");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("03 Primary Unitholder Language Preference");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("03 Primary Unitholder Language Preference");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGR02(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CIR02.pdf");

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // last name
            var lastName = formDocument.Form.FindFormField("03 Unitholder Last Name");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("03 Unitholder Info First Name");
            firstName.Value = fd.FirstName;

            // address postal
            var address = formDocument.Form.FindFormField("03 Unitholder Street Addr");
            address.Value = fd.Address;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Unitholder Postal Code");
            postalCode.Value = fd.PostalCode;

            // city
            var city = formDocument.Form.FindFormField("03 Unitholder City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Unitholder Province");
            province.Value = fd.Province;

            // cell
            var cell = formDocument.Form.FindFormField("03 Unitholder Mobile No");
            cell.Value = fd.Phone;

            // email
            var email = formDocument.Form.FindFormField("03 Unitholder Email Address");
            email.Value = fd.Email;

            // dob
            var dob = formDocument.Form.FindFormField("03 Unitholder DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("03 Unitholder SIN Bus No");
            sin.Value = fd.SIN.ToString();

            // language
            var lang = fd.Language;
            if (lang == "Français")
            {
                var fr = formDocument.Form.FindFormField("03 Primary Unitholder Language Preference");
                fr.Value = "Yes";
            }

            if (lang == "English")
            {
                var en = formDocument.Form.FindFormField("03 Primary Unitholder Language Preference");
                en.Value = "Yes";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGCE0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CICE0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("01 Dealer Number");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("01 Rep Number");
            agentCode.Value = fd.AgentCode;

            // dealer account #
            var dealerAccNum = formDocument.Form.FindFormField("01 Dealer Acct No");
            dealerAccNum.Value = DealerAccNumCIG;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("01 Securityholder - Salutation Options#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("01 Securityholder - Salutation Options#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("01 Securityholder - Salutation Options#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("01 Securityholder - Salutation Options#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // last name
            var lastName = formDocument.Form.FindFormField("02 Securityholder Info - Last Name");
            lastName.Value = fd.LastName;

            // first name
            var firstName = formDocument.Form.FindFormField("02 Securityholder Info - First Name");
            firstName.Value = fd.FirstName;

            // email
            var email = formDocument.Form.FindFormField("02 Securityholder Info - Email");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("02 Securityholder Info - Resident Address");
            address.Value = fd.Address;

            // postal code
            var postalCode = formDocument.Form.FindFormField("02 Securityholder Info - Residential Postal Code");
            postalCode.Value = fd.PostalCode;

            // city
            var city = formDocument.Form.FindFormField("02 Securityholder Info - Residential City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("02 Securityholder Info - Residential Province");
            province.Value = fd.Province;

            // cell
            var cell = formDocument.Form.FindFormField("02 Securityholder Info - Mobile Phone No");
            cell.Value = fd.Phone;

            // dob
            var dob = formDocument.Form.FindFormField("02 Securityholder Info - DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // sin
            var sin = formDocument.Form.FindFormField("02 Securityholder Info - SIN");
            sin.Value = fd.SIN;

            // language
            //switch (fd.Language)
            //{
            //    case "Français":
            //        var lang = formDocument.Form.FindFormField("01 Securityholder - Language Options#1");
            //        lang.Value = "Yes";
            //            break;
            //    case "English":
            //        lang = formDocument.Form.FindFormField("01 Securityholder - Language Options#0");
            //        lang.Value = "Yes";
            //        break;
            //}

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGF01(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CIF01.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("02 Dealer No");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("03 Unitholder Info First Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("03 Unitholder Last Name");
            lastName.Value = fd.LastName;

            // email
            var email = formDocument.Form.FindFormField("03 Unitholder Email Address");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("03 Unitholder Street Addr");
            address.Value = fd.Address;


            // city
            var city = formDocument.Form.FindFormField("03 Unitholder City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Unitholder Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Unitholder Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("03 Unitholder Mobile No");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("03 Unitholder SIN Bus No");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("03 Unitholder DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGF02(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CIF02.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("02 Dealer No");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("03 Unitholder Info First Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("03 Unitholder Last Name");
            lastName.Value = fd.LastName;

            // email
            var email = formDocument.Form.FindFormField("03 Unitholder Email Address");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("03 Unitholder Street Addr");
            address.Value = fd.Address;


            // city
            var city = formDocument.Form.FindFormField("03 Unitholder City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Unitholder Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Unitholder Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("03 Unitholder Mobile No");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("03 Unitholder SIN Bus No");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("03 Unitholder DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGCR0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CICR0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("02 Dealer No");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("03 Unitholder Info First Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("03 Unitholder Last Name");
            lastName.Value = fd.LastName;

            // email
            var email = formDocument.Form.FindFormField("03 Unitholder Email Address");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("03 Unitholder Street Addr");
            address.Value = fd.Address;


            // city
            var city = formDocument.Form.FindFormField("03 Unitholder City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Unitholder Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Unitholder Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("03 Unitholder Mobile No");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("03 Unitholder SIN Bus No");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("03 Unitholder DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGRI0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CICR0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("02 Dealer No");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("03 Unitholder Info First Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("03 Unitholder Last Name");
            lastName.Value = fd.LastName;

            // email
            var email = formDocument.Form.FindFormField("03 Unitholder Email Address");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("03 Unitholder Street Addr");
            address.Value = fd.Address;


            // city
            var city = formDocument.Form.FindFormField("03 Unitholder City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Unitholder Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Unitholder Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("03 Unitholder Mobile No");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("03 Unitholder SIN Bus No");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("03 Unitholder DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGFR0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CIFR0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("02 Dealer No");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("03 Unitholder Info First Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("03 Unitholder Last Name");
            lastName.Value = fd.LastName;

            // email
            var email = formDocument.Form.FindFormField("03 Unitholder Email Address");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("03 Unitholder Street Addr");
            address.Value = fd.Address;


            // city
            var city = formDocument.Form.FindFormField("03 Unitholder City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Unitholder Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Unitholder Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("03 Unitholder Mobile No");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("03 Unitholder SIN Bus No");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("03 Unitholder DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGFP0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CIFP0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("02 Dealer No");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("03 Unitholder Info First Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("03 Unitholder Last Name");
            lastName.Value = fd.LastName;

            // email
            var email = formDocument.Form.FindFormField("03 Unitholder Email Address");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("03 Unitholder Street Addr");
            address.Value = fd.Address;


            // city
            var city = formDocument.Form.FindFormField("03 Unitholder City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Unitholder Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Unitholder Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("03 Unitholder Mobile No");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("03 Unitholder SIN Bus No");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("03 Unitholder DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGOC0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CIOC0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("02 Dealer No");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("03 Unitholder Info First Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("03 Unitholder Last Name");
            lastName.Value = fd.LastName;

            // email
            var email = formDocument.Form.FindFormField("03 Unitholder Email Address");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("03 Unitholder Street Addr");
            address.Value = fd.Address;


            // city
            var city = formDocument.Form.FindFormField("03 Unitholder City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Unitholder Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Unitholder Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("03 Unitholder Mobile No");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("03 Unitholder SIN Bus No");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("03 Unitholder DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGOC1(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CIOC1.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("02 Dealer No");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Primary Unitholder Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("03 Unitholder Info First Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("03 Unitholder Last Name");
            lastName.Value = fd.LastName;

            // email
            var email = formDocument.Form.FindFormField("03 Unitholder Email Address");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("03 Unitholder Street Addr");
            address.Value = fd.Address;


            // city
            var city = formDocument.Form.FindFormField("03 Unitholder City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Unitholder Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Unitholder Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("03 Unitholder Mobile No");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("03 Unitholder SIN Bus No");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("03 Unitholder DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGCAP(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CICAP.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("01 Dealer Number");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("01 Rep Number");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("01 Securityholder - Salutation Options#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("01 Securityholder - Salutation Options#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("01 Securityholder - Salutation Options#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("01 Securityholder - Salutation Options#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("02 Securityholder Info - First Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("02 Securityholder Info - Last Name");
            lastName.Value = fd.LastName;

            // email
            var email = formDocument.Form.FindFormField("02 Securityholder Info - Email");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("02 Securityholder Info - Resident Address");
            address.Value = fd.Address;


            // city
            var city = formDocument.Form.FindFormField("02 Securityholder Info - Residential City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("02 Securityholder Info - Residential Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("02 Securityholder Info - Residential Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("02 Securityholder Info - Mobile Phone No");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("02 Securityholder Info - SIN");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("02 Securityholder Info - DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGREI(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CIREI.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("02 Dealer No");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Subscriber Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Subscriber Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Subscriber Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Subscriber Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("03 Subscriber - First Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("03 Subscriber - Last Name");
            lastName.Value = fd.LastName;

            // email
            var email = formDocument.Form.FindFormField("03 Subscriber - Email Address");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("03 Subscriber - Street Addr");
            address.Value = fd.Address;


            // city
            var city = formDocument.Form.FindFormField("03 Subscriber - City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Subscriber - Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Subscriber - Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("03 Subscriber - Mobile No");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("03 Subscriber - SIN");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("003 Subscriber - DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            return formDocument.Stream;
        }

        private static MemoryStream FillCIGREF(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{CigPath}CIREF.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("02 Dealer No");
            dealerCode.Value = DealerCodeCIG;

            // agent code
            var agentCode = formDocument.Form.FindFormField("02 Rep No");
            agentCode.Value = fd.AgentCode;

            // salutation
            //switch (fd.Salutation)
            //{
            //    case Constants.Mr:
            //        var salute = formDocument.Form.FindFormField("03 Subscriber Salutation#0");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Mrs:
            //        salute = formDocument.Form.FindFormField("03 Subscriber Salutation#1");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Ms:
            //        salute = formDocument.Form.FindFormField("03 Subscriber Salutation#2");
            //        salute.Value = "Yes";
            //        break;
            //    case Constants.Dr:
            //        salute = formDocument.Form.FindFormField("03 Subscriber Salutation#3");
            //        salute.Value = "Yes";
            //        break;
            //    default:
            //        break;
            //}

            // first name
            var firstName = formDocument.Form.FindFormField("03 Subscriber - First Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("03 Subscriber - Last Name");
            lastName.Value = fd.LastName;

            // email
            var email = formDocument.Form.FindFormField("03 Subscriber - Email Address");
            email.Value = fd.Email;

            // address
            var address = formDocument.Form.FindFormField("03 Subscriber - Street Addr");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("03 Subscriber - City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("03 Subscriber - Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("03 Subscriber - Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("03 Subscriber - Mobile No");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("03 Subscriber - SIN");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("003 Subscriber - DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            return formDocument.Stream;
        }

        private static MemoryStream FillFIDR01(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDR01.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("001-2");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("001-3");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("004-1");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M.";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr.";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("004-2");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("004-3");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("004-4");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("004-7");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("004-8");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("004-6b");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("004-10b");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("004-13");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("004-12");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("004-14");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("004-15");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Français";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }
        
        private static MemoryStream FillFIDR02(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDR02.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("001-2");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("001-3");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("004-1");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M.";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr.";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("004-2");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("004-3");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("004-4");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("004-7");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("004-8");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("004-6b");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("004-10b");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("004-13");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("004-12");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("004-14");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("004-15");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Français";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillFIDCE0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDCE0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("001-2");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("001-3");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("003-1");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M.";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr.";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("003-2");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("003-3");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("003-4");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("003-7");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("003-8");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("003-6b");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("003-10b");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("003-13");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("003-12");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("003-14");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("003-15");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Français";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillFIDF01(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDF01.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("002");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("003");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("016");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M.";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr.";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("018");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("017");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("027");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("029");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("030");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("031");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("026");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("021");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("020");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("025");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("022");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Français";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillFIDF02(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDF02.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("001-2");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("001-3");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("004-1");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M.";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr.";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("004-2");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("004-3");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("004-4");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("004-7");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("004-5");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("004-6b");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("004-10b");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("004-13");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("004-12");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("004-14");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("004-15");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Français";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillFIDCR0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDCR0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("001-2");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("001-3");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("004-1");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M.";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr.";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("004-2");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("004-3");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("004-4");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("004-7");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("004-8");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("004-6b");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("004-10b");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("004-13");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("004-12");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("004-14");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("004-15");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Français";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillFIDRI0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDRI0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("001-2");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("001-3");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("004-1");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M.";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr.";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("004-2");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("004-3");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("004-4");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("004-7");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("004-8");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("004-6b");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("004-10b");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("004-13");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("004-12");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("004-14");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("004-15");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Français";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillFIDFR0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDFR0.pdf");

            /// dealer code
            var dealerCode = formDocument.Form.FindFormField("002");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("003");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("016");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M.";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr.";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("018");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("017");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("027");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("029");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("030");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("031");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("026");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("021");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("020");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("025");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("022");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Français";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillFIDOC0(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDOC0.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("001-2");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("001-3");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("004-1");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M.";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr.";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("004-2");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("004-3");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("004-4");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("004-7");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("004-8");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("004-6b");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("004-10b");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("004-13");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("004-12");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("004-14");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("004-15");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Français";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillFIDCAP(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDCAP.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("IP D code");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("IP D code");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("Account Holder Salutation");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("Account Holder Name");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("Account Holder Last Name");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("Account Holder Street Address");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("Account Holder City");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("Account Holder Province");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("Account Holder Postal Code");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("Cell phone");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("Account Holder SIN");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("Account Holder DOB");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("Account Holder Email");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("Account Holder Language");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Fran?ais";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillFIDREI(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDREI.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("001-2");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("001-3");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("003-1");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M.";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("027");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("028");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("021");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("023");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("024");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("025");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("020b");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("017");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("016");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("019");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("018");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Français";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }

        private static MemoryStream FillFIDREF(FormData fd)
        {
            var formDocument = PdfDocument.FromFile($"{FidPath}FIDREF.pdf");

            // dealer code
            var dealerCode = formDocument.Form.FindFormField("001-1");
            dealerCode.Value = DealerCodeFID;

            // agent code
            var agentCode = formDocument.Form.FindFormField("001-2");
            agentCode.Value = fd.AgentCode;

            // salutation
            var salute = formDocument.Form.FindFormField("007");
            switch (fd.Salutation)
            {
                case Constants.Mr:
                    salute.Value = "M.";
                    break;
                case Constants.Mrs:
                    salute.Value = "Mme";
                    break;
                case Constants.Ms:
                    salute.Value = "Mlle";
                    break;
                case Constants.Dr:
                    salute.Value = "Dr";
                    break;
                default:
                    break;
            }

            // first name
            var firstName = formDocument.Form.FindFormField("009");
            firstName.Value = fd.FirstName;

            // last name
            var lastName = formDocument.Form.FindFormField("008");
            lastName.Value = fd.LastName;

            // address
            var address = formDocument.Form.FindFormField("018");
            address.Value = fd.Address;

            // city
            var city = formDocument.Form.FindFormField("020");
            city.Value = fd.City;

            // province
            var province = formDocument.Form.FindFormField("021");
            province.Value = fd.Province;

            // postal code
            var postalCode = formDocument.Form.FindFormField("022");
            postalCode.Value = fd.PostalCode;

            // cell
            var cell = formDocument.Form.FindFormField("017b");
            cell.Value = fd.Phone;

            // sin
            var sin = formDocument.Form.FindFormField("012");
            sin.Value = fd.SIN;

            // dob
            var dob = formDocument.Form.FindFormField("011");
            dob.Value = fd.DOB.ToString("dd/MM/yyyy");

            // email
            var email = formDocument.Form.FindFormField("016");
            email.Value = fd.Email;

            // language
            var langField = formDocument.Form.FindFormField("013");
            var lang = fd.Language;
            if (lang == "Français")
            {
                langField.Value = "Français";
            }

            if (lang == "English")
            {
                langField.Value = "Anglais";
            }

            return formDocument.Stream;
        }
    }
}


