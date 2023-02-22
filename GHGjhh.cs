
private void GenerateContract()

        {

            Word2Pdf objWorPdf = new Word2Pdf();

            GlobalProcedures.SplashScreenShow(this, typeof(WaitForms.FPrintDocumentWait));

            object fileName = Path.Combine(GlobalVariables.V_ExecutingFolder + "\\Documents\\" + GlobalVariables.V_WindowsUserName + "\\Müqavilə.docx");

            if (!File.Exists(fileName.ToString()))

            {

                GlobalProcedures.ShowWarningMessage("Müqavilənin faylı tapılmadı.");

                GlobalProcedures.SplashScreenClose();

                return;

            }

            code_number = int.Parse(Regex.Replace(RegisterCodeText.Text, "[^0-9]", ""));

            string filePath = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document aDoc = null;

            object saveAs = Path.Combine(filePath);

            object readOnly = false;

            object isVisible = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            aDoc.Activate();

            double d = 0, div = 0;

            int mod = 0;

            

            try

            {

                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$fullname]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$cardnumber]", CardNumberValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$pincode]", PinCodeValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$issue_date]",  IssueDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$reliable_date]", ReliableDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$registered_address]", RegisteredAddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$address]", AddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$phone]", PhonesText.Text);

                //if (customer_type_id == 1)

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ", " + IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib)");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib");

                //}

                //else

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ")");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", null);

                //}

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyname]", GlobalVariables.V_CompanyName);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyvoen]", GlobalVariables.V_CompanyVoen);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyphone]", GlobalVariables.V_CompanyPhone);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyaddress]", GlobalVariables.V_CompanyAddress);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companydirector]", GlobalVariables.V_CompanyDirector);

                if (File.Exists(filePath))

                    File.Delete(filePath);

                aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Close(ref missing, ref missing, ref missing);

                string strFileName = "Müqavilə.docx";

                object FromLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

                string FileExtension = Path.GetExtension(strFileName);

                string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");

                if (FileExtension == ".doc" || FileExtension == ".docx")

                {

                    object ToLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.pdf";

                    objWorPdf.InputLocation = FromLocation;

                    objWorPdf.OutputLocation = ToLocation;

                    objWorPdf.Word2PdfCOnversion();

                }

            }

            catch

            {

                GlobalProcedures.SplashScreenClose();

                GlobalProcedures.ShowErrorMessage(code_number + "_Müqavilə.docx faylı yaradılmadı.");

            }

            finally

            {

                GlobalProcedures.SplashScreenClose();

            }

        }private void GenerateContract()

        {

            Word2Pdf objWorPdf = new Word2Pdf();

            GlobalProcedures.SplashScreenShow(this, typeof(WaitForms.FPrintDocumentWait));

            object fileName = Path.Combine(GlobalVariables.V_ExecutingFolder + "\\Documents\\" + GlobalVariables.V_WindowsUserName + "\\Müqavilə.docx");

            if (!File.Exists(fileName.ToString()))

            {

                GlobalProcedures.ShowWarningMessage("Müqavilənin faylı tapılmadı.");

                GlobalProcedures.SplashScreenClose();

                return;

            }

            code_number = int.Parse(Regex.Replace(RegisterCodeText.Text, "[^0-9]", ""));

            string filePath = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document aDoc = null;

            object saveAs = Path.Combine(filePath);

            object readOnly = false;

            object isVisible = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            aDoc.Activate();

            double d = 0, div = 0;

            int mod = 0;

            

            try

            {

                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$fullname]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$cardnumber]", CardNumberValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$pincode]", PinCodeValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$issue_date]",  IssueDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$reliable_date]", ReliableDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$registered_address]", RegisteredAddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$address]", AddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$phone]", PhonesText.Text);

                //if (customer_type_id == 1)

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ", " + IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib)");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib");

                //}

                //else

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ")");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", null);

                //}

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyname]", GlobalVariables.V_CompanyName);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyvoen]", GlobalVariables.V_CompanyVoen);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyphone]", GlobalVariables.V_CompanyPhone);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyaddress]", GlobalVariables.V_CompanyAddress);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companydirector]", GlobalVariables.V_CompanyDirector);

                if (File.Exists(filePath))

                    File.Delete(filePath);

                aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Close(ref missing, ref missing, ref missing);

                string strFileName = "Müqavilə.docx";

                object FromLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

                string FileExtension = Path.GetExtension(strFileName);

                string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");

                if (FileExtension == ".doc" || FileExtension == ".docx")

                {

                    object ToLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.pdf";

                    objWorPdf.InputLocation = FromLocation;

                    objWorPdf.OutputLocation = ToLocation;

                    objWorPdf.Word2PdfCOnversion();

                }

            }

            catch

            {

                GlobalProcedures.SplashScreenClose();

                GlobalProcedures.ShowErrorMessage(code_number + "_Müqavilə.docx faylı yaradılmadı.");

            }

            finally

            {

                GlobalProcedures.SplashScreenClose();

            }

        }private void GenerateContract()

        {

            Word2Pdf objWorPdf = new Word2Pdf();

            GlobalProcedures.SplashScreenShow(this, typeof(WaitForms.FPrintDocumentWait));

            object fileName = Path.Combine(GlobalVariables.V_ExecutingFolder + "\\Documents\\" + GlobalVariables.V_WindowsUserName + "\\Müqavilə.docx");

            if (!File.Exists(fileName.ToString()))

            {

                GlobalProcedures.ShowWarningMessage("Müqavilənin faylı tapılmadı.");

                GlobalProcedures.SplashScreenClose();

                return;

            }

            code_number = int.Parse(Regex.Replace(RegisterCodeText.Text, "[^0-9]", ""));

            string filePath = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document aDoc = null;

            object saveAs = Path.Combine(filePath);

            object readOnly = false;

            object isVisible = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            aDoc.Activate();

            double d = 0, div = 0;

            int mod = 0;

            

            try

            {

                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$fullname]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$cardnumber]", CardNumberValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$pincode]", PinCodeValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$issue_date]",  IssueDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$reliable_date]", ReliableDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$registered_address]", RegisteredAddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$address]", AddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$phone]", PhonesText.Text);

                //if (customer_type_id == 1)

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ", " + IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib)");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib");

                //}

                //else

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ")");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", null);

                //}

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyname]", GlobalVariables.V_CompanyName);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyvoen]", GlobalVariables.V_CompanyVoen);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyphone]", GlobalVariables.V_CompanyPhone);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyaddress]", GlobalVariables.V_CompanyAddress);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companydirector]", GlobalVariables.V_CompanyDirector);

                if (File.Exists(filePath))

                    File.Delete(filePath);

                aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Close(ref missing, ref missing, ref missing);

                string strFileName = "Müqavilə.docx";

                object FromLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

                string FileExtension = Path.GetExtension(strFileName);

                string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");

                if (FileExtension == ".doc" || FileExtension == ".docx")

                {

                    object ToLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.pdf";

                    objWorPdf.InputLocation = FromLocation;

                    objWorPdf.OutputLocation = ToLocation;

                    objWorPdf.Word2PdfCOnversion();

                }

            }

            catch

            {

                GlobalProcedures.SplashScreenClose();

                GlobalProcedures.ShowErrorMessage(code_number + "_Müqavilə.docx faylı yaradılmadı.");

            }

            finally

            {

                GlobalProcedures.SplashScreenClose();

            }

        }










private void GenerateContract()

        {

            Word2Pdf objWorPdf = new Word2Pdf();

            GlobalProcedures.SplashScreenShow(this, typeof(WaitForms.FPrintDocumentWait));

            object fileName = Path.Combine(GlobalVariables.V_ExecutingFolder + "\\Documents\\" + GlobalVariables.V_WindowsUserName + "\\Müqavilə.docx");

            if (!File.Exists(fileName.ToString()))

            {

                GlobalProcedures.ShowWarningMessage("Müqavilənin faylı tapılmadı.");

                GlobalProcedures.SplashScreenClose();

                return;

            }

            code_number = int.Parse(Regex.Replace(RegisterCodeText.Text, "[^0-9]", ""));

            string filePath = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document aDoc = null;

            object saveAs = Path.Combine(filePath);

            object readOnly = false;

            object isVisible = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            aDoc.Activate();

            double d = 0, div = 0;

            int mod = 0;

            

            try

            {

                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$fullname]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$cardnumber]", CardNumberValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$pincode]", PinCodeValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$issue_date]",  IssueDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$reliable_date]", ReliableDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$registered_address]", RegisteredAddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$address]", AddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$phone]", PhonesText.Text);

                //if (customer_type_id == 1)

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ", " + IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib)");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib");

                //}

                //else

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ")");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", null);

                //}

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyname]", GlobalVariables.V_CompanyName);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyvoen]", GlobalVariables.V_CompanyVoen);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyphone]", GlobalVariables.V_CompanyPhone);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyaddress]", GlobalVariables.V_CompanyAddress);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companydirector]", GlobalVariables.V_CompanyDirector);

                if (File.Exists(filePath))

                    File.Delete(filePath);

                aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Close(ref missing, ref missing, ref missing);

                string strFileName = "Müqavilə.docx";

                object FromLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

                string FileExtension = Path.GetExtension(strFileName);

                string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");

                if (FileExtension == ".doc" || FileExtension == ".docx")

                {

                    object ToLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.pdf";

                    objWorPdf.InputLocation = FromLocation;

                    objWorPdf.OutputLocation = ToLocation;

                    objWorPdf.Word2PdfCOnversion();

                }

            }

            catch

            {

                GlobalProcedures.SplashScreenClose();

                GlobalProcedures.ShowErrorMessage(code_number + "_Müqavilə.docx faylı yaradılmadı.");

            }

            finally

            {

                GlobalProcedures.SplashScreenClose();

            }

        }private void GenerateContract()

        {

            Word2Pdf objWorPdf = new Word2Pdf();

            GlobalProcedures.SplashScreenShow(this, typeof(WaitForms.FPrintDocumentWait));

            object fileName = Path.Combine(GlobalVariables.V_ExecutingFolder + "\\Documents\\" + GlobalVariables.V_WindowsUserName + "\\Müqavilə.docx");

            if (!File.Exists(fileName.ToString()))

            {

                GlobalProcedures.ShowWarningMessage("Müqavilənin faylı tapılmadı.");

                GlobalProcedures.SplashScreenClose();

                return;

            }

            code_number = int.Parse(Regex.Replace(RegisterCodeText.Text, "[^0-9]", ""));

            string filePath = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document aDoc = null;

            object saveAs = Path.Combine(filePath);

            object readOnly = false;

            object isVisible = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            aDoc.Activate();

            double d = 0, div = 0;

            int mod = 0;

            

            try

            {

                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$fullname]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$cardnumber]", CardNumberValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$pincode]", PinCodeValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$issue_date]",  IssueDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$reliable_date]", ReliableDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$registered_address]", RegisteredAddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$address]", AddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$phone]", PhonesText.Text);

                //if (customer_type_id == 1)

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ", " + IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib)");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib");

                //}

                //else

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ")");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", null);

                //}

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyname]", GlobalVariables.V_CompanyName);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyvoen]", GlobalVariables.V_CompanyVoen);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyphone]", GlobalVariables.V_CompanyPhone);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyaddress]", GlobalVariables.V_CompanyAddress);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companydirector]", GlobalVariables.V_CompanyDirector);

                if (File.Exists(filePath))

                    File.Delete(filePath);

                aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Close(ref missing, ref missing, ref missing);

                string strFileName = "Müqavilə.docx";

                object FromLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

                string FileExtension = Path.GetExtension(strFileName);

                string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");

                if (FileExtension == ".doc" || FileExtension == ".docx")

                {

                    object ToLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.pdf";

                    objWorPdf.InputLocation = FromLocation;

                    objWorPdf.OutputLocation = ToLocation;

                    objWorPdf.Word2PdfCOnversion();

                }

            }

            catch

            {

                GlobalProcedures.SplashScreenClose();

                GlobalProcedures.ShowErrorMessage(code_number + "_Müqavilə.docx faylı yaradılmadı.");

            }

            finally

            {

                GlobalProcedures.SplashScreenClose();

            }

        }private void GenerateContract()

        {

            Word2Pdf objWorPdf = new Word2Pdf();

            GlobalProcedures.SplashScreenShow(this, typeof(WaitForms.FPrintDocumentWait));

            object fileName = Path.Combine(GlobalVariables.V_ExecutingFolder + "\\Documents\\" + GlobalVariables.V_WindowsUserName + "\\Müqavilə.docx");

            if (!File.Exists(fileName.ToString()))

            {

                GlobalProcedures.ShowWarningMessage("Müqavilənin faylı tapılmadı.");

                GlobalProcedures.SplashScreenClose();

                return;

            }

            code_number = int.Parse(Regex.Replace(RegisterCodeText.Text, "[^0-9]", ""));

            string filePath = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document aDoc = null;

            object saveAs = Path.Combine(filePath);

            object readOnly = false;

            object isVisible = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            aDoc.Activate();

            double d = 0, div = 0;

            int mod = 0;

            

            try

            {

                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$fullname]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$cardnumber]", CardNumberValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$pincode]", PinCodeValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$issue_date]",  IssueDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$reliable_date]", ReliableDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$registered_address]", RegisteredAddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$address]", AddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$phone]", PhonesText.Text);

                //if (customer_type_id == 1)

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ", " + IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib)");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib");

                //}

                //else

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ")");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", null);

                //}

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyname]", GlobalVariables.V_CompanyName);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyvoen]", GlobalVariables.V_CompanyVoen);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyphone]", GlobalVariables.V_CompanyPhone);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyaddress]", GlobalVariables.V_CompanyAddress);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companydirector]", GlobalVariables.V_CompanyDirector);

                if (File.Exists(filePath))

                    File.Delete(filePath);

                aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Close(ref missing, ref missing, ref missing);

                string strFileName = "Müqavilə.docx";

                object FromLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

                string FileExtension = Path.GetExtension(strFileName);

                string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");

                if (FileExtension == ".doc" || FileExtension == ".docx")

                {

                    object ToLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.pdf";

                    objWorPdf.InputLocation = FromLocation;

                    objWorPdf.OutputLocation = ToLocation;

                    objWorPdf.Word2PdfCOnversion();

                }

            }

            catch

            {

                GlobalProcedures.SplashScreenClose();

                GlobalProcedures.ShowErrorMessage(code_number + "_Müqavilə.docx faylı yaradılmadı.");

            }

            finally

            {

                GlobalProcedures.SplashScreenClose();

            }

        }private void GenerateContract()

        {

            Word2Pdf objWorPdf = new Word2Pdf();

            GlobalProcedures.SplashScreenShow(this, typeof(WaitForms.FPrintDocumentWait));

            object fileName = Path.Combine(GlobalVariables.V_ExecutingFolder + "\\Documents\\" + GlobalVariables.V_WindowsUserName + "\\Müqavilə.docx");

            if (!File.Exists(fileName.ToString()))

            {

                GlobalProcedures.ShowWarningMessage("Müqavilənin faylı tapılmadı.");

                GlobalProcedures.SplashScreenClose();

                return;

            }

            code_number = int.Parse(Regex.Replace(RegisterCodeText.Text, "[^0-9]", ""));

            string filePath = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document aDoc = null;

            object saveAs = Path.Combine(filePath);

            object readOnly = false;

            object isVisible = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            aDoc.Activate();

            double d = 0, div = 0;

            int mod = 0;

            

            try

            {

                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$fullname]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$cardnumber]", CardNumberValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$pincode]", PinCodeValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$issue_date]",  IssueDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$reliable_date]", ReliableDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$registered_address]", RegisteredAddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$address]", AddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$phone]", PhonesText.Text);

                //if (customer_type_id == 1)

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ", " + IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib)");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib");

                //}

                //else

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ")");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", null);

                //}

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyname]", GlobalVariables.V_CompanyName);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyvoen]", GlobalVariables.V_CompanyVoen);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyphone]", GlobalVariables.V_CompanyPhone);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyaddress]", GlobalVariables.V_CompanyAddress);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companydirector]", GlobalVariables.V_CompanyDirector);

                if (File.Exists(filePath))

                    File.Delete(filePath);

                aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Close(ref missing, ref missing, ref missing);

                string strFileName = "Müqavilə.docx";

                object FromLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

                string FileExtension = Path.GetExtension(strFileName);

                string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");

                if (FileExtension == ".doc" || FileExtension == ".docx")

                {

                    object ToLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.pdf";

                    objWorPdf.InputLocation = FromLocation;

                    objWorPdf.OutputLocation = ToLocation;

                    objWorPdf.Word2PdfCOnversion();

                }

            }

            catch

            {

                GlobalProcedures.SplashScreenClose();

                GlobalProcedures.ShowErrorMessage(code_number + "_Müqavilə.docx faylı yaradılmadı.");

            }

            finally

            {

                GlobalProcedures.SplashScreenClose();

            }

        }private void GenerateContract()

        {

            Word2Pdf objWorPdf = new Word2Pdf();

            GlobalProcedures.SplashScreenShow(this, typeof(WaitForms.FPrintDocumentWait));

            object fileName = Path.Combine(GlobalVariables.V_ExecutingFolder + "\\Documents\\" + GlobalVariables.V_WindowsUserName + "\\Müqavilə.docx");

            if (!File.Exists(fileName.ToString()))

            {

                GlobalProcedures.ShowWarningMessage("Müqavilənin faylı tapılmadı.");

                GlobalProcedures.SplashScreenClose();

                return;

            }

            code_number = int.Parse(Regex.Replace(RegisterCodeText.Text, "[^0-9]", ""));

            string filePath = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document aDoc = null;

            object saveAs = Path.Combine(filePath);

            object readOnly = false;

            object isVisible = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            aDoc.Activate();

            double d = 0, div = 0;

            int mod = 0;

            

            try

            {

                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$fullname]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$cardnumber]", CardNumberValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$pincode]", PinCodeValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$issue_date]",  IssueDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$reliable_date]", ReliableDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$registered_address]", RegisteredAddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$address]", AddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$phone]", PhonesText.Text);

                //if (customer_type_id == 1)

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ", " + IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib)");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib");

                //}

                //else

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ")");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", null);

                //}

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyname]", GlobalVariables.V_CompanyName);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyvoen]", GlobalVariables.V_CompanyVoen);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyphone]", GlobalVariables.V_CompanyPhone);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyaddress]", GlobalVariables.V_CompanyAddress);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companydirector]", GlobalVariables.V_CompanyDirector);

                if (File.Exists(filePath))

                    File.Delete(filePath);

                aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Close(ref missing, ref missing, ref missing);

                string strFileName = "Müqavilə.docx";

                object FromLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

                string FileExtension = Path.GetExtension(strFileName);

                string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");

                if (FileExtension == ".doc" || FileExtension == ".docx")

                {

                    object ToLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.pdf";

                    objWorPdf.InputLocation = FromLocation;

                    objWorPdf.OutputLocation = ToLocation;

                    objWorPdf.Word2PdfCOnversion();

                }

            }

            catch

            {

                GlobalProcedures.SplashScreenClose();

                GlobalProcedures.ShowErrorMessage(code_number + "_Müqavilə.docx faylı yaradılmadı.");

            }

            finally

            {

                GlobalProcedures.SplashScreenClose();

            }

        }private void GenerateContract()

        {

            Word2Pdf objWorPdf = new Word2Pdf();

            GlobalProcedures.SplashScreenShow(this, typeof(WaitForms.FPrintDocumentWait));

            object fileName = Path.Combine(GlobalVariables.V_ExecutingFolder + "\\Documents\\" + GlobalVariables.V_WindowsUserName + "\\Müqavilə.docx");

            if (!File.Exists(fileName.ToString()))

            {

                GlobalProcedures.ShowWarningMessage("Müqavilənin faylı tapılmadı.");

                GlobalProcedures.SplashScreenClose();

                return;

            }

            code_number = int.Parse(Regex.Replace(RegisterCodeText.Text, "[^0-9]", ""));

            string filePath = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document aDoc = null;

            object saveAs = Path.Combine(filePath);

            object readOnly = false;

            object isVisible = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            aDoc.Activate();

            double d = 0, div = 0;

            int mod = 0;

            

            try

            {

                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$fullname]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$cardnumber]", CardNumberValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$pincode]", PinCodeValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$issue_date]",  IssueDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$reliable_date]", ReliableDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$registered_address]", RegisteredAddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$address]", AddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$phone]", PhonesText.Text);

                //if (customer_type_id == 1)

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ", " + IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib)");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib");

                //}

                //else

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ")");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", null);

                //}

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyname]", GlobalVariables.V_CompanyName);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyvoen]", GlobalVariables.V_CompanyVoen);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyphone]", GlobalVariables.V_CompanyPhone);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyaddress]", GlobalVariables.V_CompanyAddress);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companydirector]", GlobalVariables.V_CompanyDirector);

                if (File.Exists(filePath))

                    File.Delete(filePath);

                aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Close(ref missing, ref missing, ref missing);

                string strFileName = "Müqavilə.docx";

                object FromLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

                string FileExtension = Path.GetExtension(strFileName);

                string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");

                if (FileExtension == ".doc" || FileExtension == ".docx")

                {

                    object ToLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.pdf";

                    objWorPdf.InputLocation = FromLocation;

                    objWorPdf.OutputLocation = ToLocation;

                    objWorPdf.Word2PdfCOnversion();

                }

            }

            catch

            {

                GlobalProcedures.SplashScreenClose();

                GlobalProcedures.ShowErrorMessage(code_number + "_Müqavilə.docx faylı yaradılmadı.");

            }

            finally

            {

                GlobalProcedures.SplashScreenClose();

            }

        }private void GenerateContract()

        {

            Word2Pdf objWorPdf = new Word2Pdf();

            GlobalProcedures.SplashScreenShow(this, typeof(WaitForms.FPrintDocumentWait));

            object fileName = Path.Combine(GlobalVariables.V_ExecutingFolder + "\\Documents\\" + GlobalVariables.V_WindowsUserName + "\\Müqavilə.docx");

            if (!File.Exists(fileName.ToString()))

            {

                GlobalProcedures.ShowWarningMessage("Müqavilənin faylı tapılmadı.");

                GlobalProcedures.SplashScreenClose();

                return;

            }

            code_number = int.Parse(Regex.Replace(RegisterCodeText.Text, "[^0-9]", ""));

            string filePath = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document aDoc = null;

            object saveAs = Path.Combine(filePath);

            object readOnly = false;

            object isVisible = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            aDoc.Activate();

            double d = 0, div = 0;

            int mod = 0;

            

            try

            {

                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$fullname]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$cardnumber]", CardNumberValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$pincode]", PinCodeValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$issue_date]",  IssueDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$reliable_date]", ReliableDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$registered_address]", RegisteredAddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$address]", AddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$phone]", PhonesText.Text);

                //if (customer_type_id == 1)

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ", " + IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib)");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib");

                //}

                //else

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ")");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", null);

                //}

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyname]", GlobalVariables.V_CompanyName);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyvoen]", GlobalVariables.V_CompanyVoen);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyphone]", GlobalVariables.V_CompanyPhone);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyaddress]", GlobalVariables.V_CompanyAddress);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companydirector]", GlobalVariables.V_CompanyDirector);

                if (File.Exists(filePath))

                    File.Delete(filePath);

                aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Close(ref missing, ref missing, ref missing);

                string strFileName = "Müqavilə.docx";

                object FromLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

                string FileExtension = Path.GetExtension(strFileName);

                string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");

                if (FileExtension == ".doc" || FileExtension == ".docx")

                {

                    object ToLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.pdf";

                    objWorPdf.InputLocation = FromLocation;

                    objWorPdf.OutputLocation = ToLocation;

                    objWorPdf.Word2PdfCOnversion();

                }

            }

            catch

            {

                GlobalProcedures.SplashScreenClose();

                GlobalProcedures.ShowErrorMessage(code_number + "_Müqavilə.docx faylı yaradılmadı.");

            }

            finally

            {

                GlobalProcedures.SplashScreenClose();

            }

        }private void GenerateContract()

        {

            Word2Pdf objWorPdf = new Word2Pdf();

            GlobalProcedures.SplashScreenShow(this, typeof(WaitForms.FPrintDocumentWait));

            object fileName = Path.Combine(GlobalVariables.V_ExecutingFolder + "\\Documents\\" + GlobalVariables.V_WindowsUserName + "\\Müqavilə.docx");

            if (!File.Exists(fileName.ToString()))

            {

                GlobalProcedures.ShowWarningMessage("Müqavilənin faylı tapılmadı.");

                GlobalProcedures.SplashScreenClose();

                return;

            }

            code_number = int.Parse(Regex.Replace(RegisterCodeText.Text, "[^0-9]", ""));

            string filePath = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

            object missing = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document aDoc = null;

            object saveAs = Path.Combine(filePath);

            object readOnly = false;

            object isVisible = false;

            wordApp.Visible = false;

            aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

            aDoc.Activate();

            double d = 0, div = 0;

            int mod = 0;

            

            try

            {

                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$fullname]", NameText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$cardnumber]", CardNumberValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$pincode]", PinCodeValue.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$issue_date]",  IssueDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$reliable_date]", ReliableDateText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$registered_address]", RegisteredAddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$address]", AddressText.Text);

                GlobalProcedures.FindAndReplace(wordApp, "[$phone]", PhonesText.Text);

                //if (customer_type_id == 1)

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ", " + IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib)");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", IssuingDateText.Text + " tarixində " + IssuingText.Text + " tərəfindən verilib");

                //}

                //else

                //{

                //    GlobalProcedures.FindAndReplace(wordApp, "[$customer]", CustomerFullNameText.Text + " (" + CardDescriptionText.Text + ")");

                //    GlobalProcedures.FindAndReplace(wordApp, "[$carddate]", null);

                //}

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyname]", GlobalVariables.V_CompanyName);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyvoen]", GlobalVariables.V_CompanyVoen);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyphone]", GlobalVariables.V_CompanyPhone);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companyaddress]", GlobalVariables.V_CompanyAddress);

                //GlobalProcedures.FindAndReplace(wordApp, "[$companydirector]", GlobalVariables.V_CompanyDirector);

                if (File.Exists(filePath))

                    File.Delete(filePath);

                aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Close(ref missing, ref missing, ref missing);

                string strFileName = "Müqavilə.docx";

                object FromLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.docx";

                string FileExtension = Path.GetExtension(strFileName);

                string ChangeExtension = strFileName.Replace(FileExtension, ".pdf");

                if (FileExtension == ".doc" || FileExtension == ".docx")

                {

                    object ToLocation = GlobalVariables.V_ExecutingFolder + "\\TEMP\\Documents\\" + code_number + "_Müqavilə.pdf";

                    objWorPdf.InputLocation = FromLocation;

                    objWorPdf.OutputLocation = ToLocation;

                    objWorPdf.Word2PdfCOnversion();

                }

            }

            catch

            {

                GlobalProcedures.SplashScreenClose();

                GlobalProcedures.ShowErrorMessage(code_number + "_Müqavilə.docx faylı yaradılmadı.");

            }

            finally

            {

                GlobalProcedures.SplashScreenClose();

            }

        }
