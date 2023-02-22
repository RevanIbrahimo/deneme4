
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
            //d = (double)CreditAmountValue.Value * 100;

            //div = (int)(d / 100);
            //mod = (int)(d % 100);
            //if (mod > 0)
            //{
            //    if (credit_currency_id == 1)
            //        qep = " " + mod.ToString() + " qəpik";
            //    else
            //        qep = " " + mod.ToString();
            //}

            //amount_with_word = "(" + GlobalFunctions.IntegerToWritten(div) + ") " + credit_currency_name + qep;

            ////Komissiya
            //d = (double)CommissionValue.Value * 100;

            //div = (int)(d / 100);
            //mod = (int)(d % 100);
            //if (mod > 0)
            //{
            //    if (credit_currency_id == 1)
            //        qep = " " + mod.ToString() + " qəpik";
            //    else
            //        qep = " " + mod.ToString();
            //}

            //com_with_word = "(" + GlobalFunctions.IntegerToWritten(div) + ") " + credit_currency_name + qep;

            ////FIFD
            //decimal fifd = Math.Round(FifdValue.Value, 2);
            //d = (double)fifd * 100;

            //div = (int)(d / 100);
            //mod = (int)(d % 100);
            //if (mod > 0)
            //    qep = " tam yüzdə " + GlobalFunctions.IntegerToWritten(mod);

            //fifd_with_word = "(" + GlobalFunctions.IntegerToWritten(div) + qep + ")";

            //if (PeriodCheckEdit.Checked)
            //    period = ContractEndDate.Text + " tarixinə qədər";
            //else
            //    period = PeriodValue.Value + " (" + GlobalFunctions.IntegerToWritten((int)PeriodValue.Value) + ") ay";

            //phone = GlobalFunctions.GetName($@"SELECT PHONE FROM CRS_USER.V_PHONE WHERE OWNER_TYPE = '{person_description}' AND OWNER_ID = {CustomerID}");

            try
            {
                GlobalProcedures.FindAndReplace(wordApp, "[$contractcode]", RegisterCodeText.Text);
                GlobalProcedures.FindAndReplace(wordApp, "[$contractdate]", OrderDateText.Text);
                GlobalProcedures.FindAndReplace(wordApp, "[$customername]", NameText.Text);
                GlobalProcedures.FindAndReplace(wordApp, "[$customerpincode]", FinCodeSearch.Text);
                GlobalProcedures.FindAndReplace(wordApp, "[$amount]", OrderAmountValue.Value.ToString());
                GlobalProcedures.FindAndReplace(wordApp, "[$firstpayment]", FirstPaymentValue.Text);
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
