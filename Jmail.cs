using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenPop.Mime;
using OpenPop.Mime.Header;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using OpenPop.Common.Logging;
using Message = OpenPop.Mime.Message;
using MySql.Data.MySqlClient;  //Its for MySQL 
//using Microsoft.Office.Interop.Excel;



namespace Jmail
{


    public partial class Jmail : Form
    {

        //string popServer = "pop3.mweb.co.za";
        string popServer ;
        int port = 110;
        bool useSsl = false;
        string login = "";
        string password = "";
        
        int count;
        StringBuilder sb = new StringBuilder();
        StringBuilder sbdb = new StringBuilder();
        string MailFrom;
        string MailDate;
        int RowCount = 3;
        string IDforDB;
        string DFBMes;
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        private readonly Pop3Client pop3Client;
        public Jmail()
        {

            InitializeComponent();
            pop3Client = new Pop3Client();


        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        private void Startbutton_Click(object sender, EventArgs e)
        {
            login = "johnny@e-c-p.co.za";
            password = "x293vgtd";
            popServer = "pop3.mweb.co.za";
            textBoxMailName.Text = "johnny@e-c-p.co.za";
            ReceiveMails();
            login = "nucrm@ecptest.co.za";
            password = "nucrmmail";
            popServer = "mail.ecptest.co.za";
            textBoxMailName.Text = "nucrm@ecptest.co.za";
            ReceiveMails();
        }

        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        //EXIT
        private void button1_Click(object sender, EventArgs e)
        {
            pop3Client.Disconnect();
            Application.Exit();

        }
        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        private void ReceiveMails()
        {
            int success = 0;
            int fail = 0;

            string MailSubject;
            byte[] MailBodyB;
            string MailBody;
            string MailBodyString;
            int index1;
            int index2;
            string EmailAddress;
            string PersonName;

            // Disable buttons while working
            Startbutton.Enabled = false;
            try
            {
                if (pop3Client.Connected)
                    pop3Client.Disconnect();
                pop3Client.Connect(popServer, port, useSsl);
                pop3Client.Authenticate(login, password);
                count = pop3Client.GetMessageCount();


                for (int i = count; i >= 1; i -= 1)
                {
                    // Check if the form is closed while we are working. If so, abort
                    if (IsDisposed)
                        return;

                    // Refresh the form while fetching emails
                    // This will fix the "Application is not responding" problem
                    Application.DoEvents();
                    Message message = pop3Client.GetMessage(i);

                    try
                    {

                        MestextBox.Text = i.ToString();
                        // exstract email address
                        MailFrom = message.Headers.From.ToString();
                        index1 = MailFrom.IndexOf("<");
                        index1++;
                        index2 = MailFrom.IndexOf(">");
                        EmailAddress = MailFrom.Substring(index1, index2 - index1);

                        //extract person name
                        index1 = EmailAddress.IndexOf("@");
                        PersonName = EmailAddress.Substring(0, index1);
                        // extract date
                        MailDate = message.Headers.Date;

                        string shortedate = MailDate.Substring(0, 25);

                        IDforDB = PersonName + shortedate;
                        IDforDB = IDforDB.Replace(",", "");
                        IDforDB = IDforDB.Replace(":", "");
                        IDforDB = IDforDB.Replace(" ", "");
                        IDforDB = IDforDB.Replace(".", "");

                        MailSubject = message.Headers.Subject;

                        if (MailSubject == "Reporter" || MailSubject == "Feedback" || MailSubject == "Consignment" || MailSubject == "Strategy")
                        {
                            textBoxID.Text = IDforDB;
                            FromtextBox.Text = EmailAddress;
                            DatetextBox.Text = shortedate;
                            SubjecttextBox.Text = MailSubject;
                            MailBody = message.MessagePart.GetBodyAsText();
                            textBoxBody.Text = MailBody;
                            ExstarctBody(MailBody);
                            index1 = 0;
                            pop3Client.DeleteMessage(i);
                        }

                        
                        success++;
                    }
                    catch (Exception e)
                    {

                        MailBodyB = message.MessagePart.MessageParts[0].Body;
                        //convert byte[] to string
                        MailBodyString = Encoding.UTF8.GetString(MailBodyB, 0, MailBodyB.Length);
                        textBoxBody.Text = MailBodyString;
                        ExstarctBody(MailBodyString);
                        index1 = 0;
                        //DefaultLogger.Log.LogError(
                        //    "TestForm: Message fetching failed: " + e.Message + "\r\n" +
                        //    "Stack trace:\r\n" +
                        //    e.StackTrace);
                        //fail++;
                        pop3Client.DeleteMessage(i);
                    }

                }

                //MessageBox.Show(this, "Mail received!\nSuccesses: " + success + "\nFailed: " + fail, "Message fetching done");

                //if (fail > 0)
                //{
                //    MessageBox.Show(this,
                //                    "Since some of the emails were not parsed correctly (exceptions were thrown)\r\n" +
                //                    "please consider sending your log file to the developer for fixing.\r\n" +
                //                    "If you are able to include any extra information, please do so.",
                //                    "Help improve OpenPop!");
                //}
            }
            catch (InvalidLoginException)
            {
                MessageBox.Show(this, "The server did not accept the user credentials!", "POP3 Server Authentication");
            }
            catch (PopServerNotFoundException)
            {
                MessageBox.Show(this, "The server could not be found", "POP3 Retrieval");
            }
            catch (PopServerLockedException)
            {
                MessageBox.Show(this, "The mailbox is locked. It might be in use or under maintenance. Are you connected elsewhere?", "POP3 Account Locked");
            }
            catch (LoginDelayException)
            {
                MessageBox.Show(this, "Login not allowed. Server enforces delay between logins. Have you connected recently?", "POP3 Account Login Delay");
            }
            catch (Exception e)
            {
                MessageBox.Show(this, "Error occurred retrieving mail. " + e.Message, "POP3 Retrieval");
            }
            finally
            {
                // Enable the buttons again
                Startbutton.Enabled = true;

            }

            
            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            void ExstarctBody(string Body)
            {
                int Irepname = 0;
                int ICompanyName = 0;
                int Iplace = 0;
                int Idebtorcode = 0;
                int Idebtorname = 0;
                int Icustomercode = 0;
                int Icustomername = 0;
                int Ispeciality = 0;
                int Ievents = 0;
                int Ipharm = 0;
                int Irepname2 = 0;
                int Ibusniness = 0;
                int Iproducts = 0;
                int Ireporttype = 0;
                int Ireportstrat = 0;
                int Ireportreport = 0;
                int Ireportfollowup = 0;
                int InotesDebtor = 0;
                int InotesCustomer = 0;
                int IstratCount = 0;
                int Irepcount = 0;
                int Ifollowcount = 0;
                int Itimestamp = 0;
                int Ilocation = 0;
                int Ilat = 0;
                int Ilong = 0;
                int ISW = 0;
                int IDV = 0;
                int Idepartment = 0;

                char[] delimiterChars = { ';', ':', '\n', '\r', '~', '=' };

                string rdata = Body;
                string[] words = rdata.Split(delimiterChars);

                string Wcode;
                int len = words.Count();
                for (int i = 0; i < len; i++)
                {

                    Wcode = words[i].ToString();

                    switch (Wcode)
                    {

                        case "Rep Name":
                            Irepname = i + 1;
                            break;
                        case "Company":
                            ICompanyName = i + 1;
                            break;
                        case "Place":
                            Iplace = i + 1;
                            break;
                        case "Debtor Code":
                            Idebtorcode = i + 1;
                            Idebtorname = i + 2;
                            break;
                        case "Customer Code":
                            Icustomercode = i + 1;
                            Icustomername = i + 2;
                            break;
                        case "Speciality":
                            Ispeciality = i + 1;
                            break;
                        case "Department":
                            Idepartment = i + 1;
                            break;
                        case "Contact":
                            Ipharm = i + 1;
                            break;
                        case "Pharmacy":
                            Ipharm = i + 1;
                            break;
                        case "Sales Person":
                            Irepname2 = i + 1;
                            break;
                        case "Business Unit":
                            Ibusniness = i + 1;
                            break;
                        case "Product":
                            Iproducts = i + 1;
                            break;
                        case "Events":
                            Ievents = i + 1;
                            break;
                        case "Report Type":
                            Ireporttype = i + 1;
                            break;
                        case "Report Strategy":
                            Ireportstrat = i + 1;
                            break;
                        case "Report Report":
                            Ireportreport = i + 1;
                            break;
                        case "Report Followup":
                            Ireportfollowup = i + 1;
                            break;
                        case "Notes Debtor":
                            InotesDebtor = i + 1;
                            break;
                        case "Notes Customer":
                            InotesCustomer = i + 1;
                            break;
                        case "LineCountStart":
                            IstratCount = i + 1;
                            break;
                        case "LineCountreport":
                            Irepcount = i + 1;
                            break;
                        case "LineCountFollow":
                            Ifollowcount = i + 1;
                            break;
                        case "Time Stamp  ":
                            Itimestamp = i + 1;
                            break;
                        case "Geo Pos  ":
                            Ilocation = i + 1;
                            break;
                        case "Latitude  ":
                            Ilat = i + 1;
                            break;
                        case "Longitude  ":
                            Ilong = i + 1;
                            break;
                        case "SW ver ":
                            ISW = i + 1;
                            break;
                        case "Data ver ":
                            IDV = i + 1;
                            break;
                        default:
                            break;
                    }
                }

                textBoxRep.Text = words[Irepname].ToString();
                textBoxCompany.Text = words[ICompanyName].ToString();
                textBoxPlace.Text = words[Iplace].ToString();
                textBoxDebtorCode.Text = words[Idebtorcode].ToString();
                if (textBoxDebtorCode.Text == "N/A")
                    textBoxDebtorName.Text = "N/A";
                else
                {
                    textBoxDebtorCode.Text = words[Idebtorcode].ToString();
                    textBoxDebtorName.Text = words[Idebtorname].ToString();
                }

                textBoxCustomerCode.Text = words[Icustomercode].ToString();
                if (textBoxCustomerCode.Text == "N/A")
                {
                    textBoxCustomerName.Text = "N/A";
                    textBoxSpeciality.Text = "N/A";
                }
                else
                {
                    textBoxCustomerCode.Text = words[Icustomercode].ToString();
                    textBoxCustomerName.Text = words[Icustomername].ToString();
                    textBoxSpeciality.Text = words[Ispeciality].ToString();
                }
                textBoxDepartment.Text = words[Idepartment].ToString();
                textBoxPharmacy.Text = words[Ipharm].ToString();
                textBoxRep2.Text = words[Irepname2].ToString();
                textBoxBussinesUnit.Text = words[Ibusniness].ToString();

                if (words[Ievents].ToString() == "Rep Name")
                    textBoxEvents.Text = " ";
                else
                    textBoxEvents.Text = words[Ievents].ToString();
                textBoxReportType.Text = words[Ireporttype].ToString();
                textBoxProduct.Text = words[Iproducts].ToString();
                textBoxObjective.Text = words[Ireportstrat].ToString();
                textBoxOutcome.Text = words[Ireportreport].ToString();
                textBoxPOA.Text = words[Ireportfollowup].ToString();
                textBoxnotesDebtor.Text = words[InotesDebtor].ToString();
                textBoxnotesCustomer.Text = words[InotesCustomer].ToString();
                textBoxTimeStamp.Text = words[Itimestamp].ToString() + ":" + words[Itimestamp + 1].ToString() + ":" + words[Itimestamp + 2].ToString();
                textBoxLocation.Text = words[Ilocation].ToString();
                textBoxLat.Text = words[Ilat].ToString();
                textBoxLong.Text = words[Ilong].ToString();
                textBoxSWver.Text = words[ISW].ToString();
                if (IDV == 0)
                    textBoxDataVer.Text = "N/A";
                else
                    textBoxDataVer.Text = words[IDV].ToString();

                //WriteToExecl();
                sbdb.Clear();
                sbdb.Append(IDforDB);
                sbdb.Append(MailFrom);
                sbdb.Append(MailDate);
                //sb.Append(textBoxTimeStamp.Text);
                //sb.Append(textBoxRep.Text);
                //sb.Append(textBoxCompany.Text);
                //sb.Append(textBoxPlace.Text);
                //sb.Append(textBoxDebtorCode.Text);
                //sb.Append(textBoxDebtorName.Text);
                //sb.Append(textBoxCustomerCode.Text);
                //sb.Append(textBoxCustomerName.Text);
                //sb.Append(textBoxSpeciality.Text);
                //sb.Append(textBoxPharmacy.Text);
                //sb.Append(textBoxRep2.Text);
                //sb.Append(textBoxBussinesUnit.Text);
                //sb.Append(textBoxEvents.Text);
                //sb.Append(textBoxReportType.Text);
                //sb.Append(textBoxProduct.Text);
                //sb.Append(textBoxObjective.Text);
                //sb.Append(textBoxOutcome.Text);
                //sb.Append(textBoxPOA.Text);
                //sb.Append(textBoxnotesDebtor.Text);
                //sb.Append(textBoxnotesCustomer.Text);
                //sb.Append(textBoxLocation.Text);
                //sb.Append(textBoxLat.Text);
                //sb.Append(textBoxLong.Text);
                //sb.Append(textBoxSWver.Text);
                WriteToDB();
            }
            //=====================================================================
            void WriteToExecl()
            {
                sb.Clear();
                sb.Append(IDforDB);
                sb.Append(MailFrom);
                sb.Append(MailDate);
                sb.Append(textBoxTimeStamp.Text);
                sb.Append(textBoxRep.Text);
                sb.Append(textBoxCompany.Text);
                sb.Append(textBoxPlace.Text);
                sb.Append(textBoxDebtorCode.Text);
                sb.Append(textBoxDebtorName.Text);
                sb.Append(textBoxCustomerCode.Text);
                sb.Append(textBoxCustomerName.Text);
                sb.Append(textBoxSpeciality.Text);
                sb.Append(textBoxPharmacy.Text);
                sb.Append(textBoxRep2.Text);
                sb.Append(textBoxBussinesUnit.Text);
                sb.Append(textBoxEvents.Text);
                sb.Append(textBoxReportType.Text);
                sb.Append(textBoxProduct.Text);
                sb.Append(textBoxObjective.Text);
                sb.Append(textBoxOutcome.Text);
                sb.Append(textBoxPOA.Text);
                sb.Append(textBoxnotesDebtor.Text);
                sb.Append(textBoxnotesCustomer.Text);
                sb.Append(textBoxLocation.Text);
                sb.Append(textBoxLat.Text);
                sb.Append(textBoxLong.Text);
                sb.Append(textBoxSWver.Text);
                sb.Append(textBoxDataVer.Text);
                Microsoft.Office.Interop.Excel.Application oXL;
                Microsoft.Office.Interop.Excel._Workbook oWB;
                Microsoft.Office.Interop.Excel._Worksheet oSheet;
                Microsoft.Office.Interop.Excel.Range oRng;
                object misvalue = System.Reflection.Missing.Value;

                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.

                oWB = oXL.Workbooks.Open("c:\\data\\reportermail.xlsx");
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                oSheet.Cells[RowCount, 1] = textBoxID.Text;
                oSheet.Cells[RowCount, 2] = textBoxTimeStamp.Text;
                oSheet.Cells[RowCount, 3] = textBoxRep.Text;
                oSheet.Cells[RowCount, 4] = textBoxCompany.Text;
                oSheet.Cells[RowCount, 5] = textBoxPlace.Text;
                oSheet.Cells[RowCount, 6] = textBoxDebtorCode.Text;
                oSheet.Cells[RowCount, 7] = textBoxDebtorName.Text;
                oSheet.Cells[RowCount, 8] = textBoxCustomerCode.Text;
                oSheet.Cells[RowCount, 9] = textBoxCustomerName.Text;
                oSheet.Cells[RowCount, 10] = textBoxSpeciality.Text;
                oSheet.Cells[RowCount, 11] = textBoxPharmacy.Text;
                oSheet.Cells[RowCount, 12] = textBoxRep2.Text;
                oSheet.Cells[RowCount, 13] = textBoxBussinesUnit.Text;
                oSheet.Cells[RowCount, 14] = textBoxEvents.Text;
                oSheet.Cells[RowCount, 15] = textBoxReportType.Text;
                oSheet.Cells[RowCount, 16] = textBoxProduct.Text;
                oSheet.Cells[RowCount, 17] = textBoxObjective.Text;
                oSheet.Cells[RowCount, 18] = textBoxOutcome.Text;
                oSheet.Cells[RowCount, 19] = textBoxPOA.Text;
                oSheet.Cells[RowCount, 20] = textBoxnotesDebtor.Text;
                oSheet.Cells[RowCount, 21] = textBoxnotesCustomer.Text;
                oSheet.Cells[RowCount, 22] = textBoxLocation.Text;
                oSheet.Cells[RowCount, 23] = textBoxLat.Text;
                oSheet.Cells[RowCount, 24] = textBoxLong.Text;
                oSheet.Cells[RowCount, 25] = textBoxSWver.Text;
                //oRng.EntireColumn.AutoFit();

                oXL.Visible = false;
                oXL.UserControl = false;
                oWB.Save();

                oWB.Close();
                RowCount++;

            }
            //====================================
            void WriteToDB()
            {
                string connetionString = null;
                if (textBoxRep.Text == "")
                    textBoxRep.Text = FromtextBox.Text;

                if (textBoxDebtorCode.Text == "")
                {
                    textBoxDebtorCode.Text = "N/A";
                    textBoxDebtorName.Text = "N/A";
                }
                if(this.textBoxCustomerCode.Text == "")
                {
                    this.textBoxCustomerCode.Text = "N/A";
                    this.textBoxCustomerName.Text = "N/A";
                    this.textBoxSpeciality.Text = "N/A";
                }
                if (this.textBoxEvents.Text == "")
                    this.textBoxEvents.Text = "N/A";
                if (textBoxLocation.Text == " " || textBoxLocation.Text == "")
                    this.textBoxLocation.Text = "NA";
                if (this.textBoxLat.Text == "" || this.textBoxLat.Text == " ")
                    this.textBoxLat.Text = "N/A";
                if (this.textBoxLong.Text == "" || this.textBoxLong.Text == " ")
                    this.textBoxLong.Text = "N/A";
                if (this.textBoxPharmacy.Text == "")
                    this.textBoxPharmacy.Text = "N/A";
                if (this.textBoxDepartment.Text == "")
                    this.textBoxDepartment.Text = "";
                if (this.textBoxObjective.Text == "")
                    this.textBoxObjective.Text = "N/A";
                if (this.textBoxOutcome.Text == "")
                    this.textBoxOutcome.Text = "N/A";
                if (this.textBoxPOA.Text == "")
                    this.textBoxPOA.Text = "N/A";
                if (this.textBoxnotesDebtor.Text == "")
                    this.textBoxnotesDebtor.Text = "N/A";
                if (this.textBoxnotesCustomer.Text == "")
                    this.textBoxnotesCustomer.Text = "N/A";
                if (this.textBoxDataVer.Text == "")
                    this.textBoxDataVer.Text = "N/A";
                

                MySqlConnection cnn;
                connetionString = "server=localhost;database=nucrm;uid=root;";
                cnn = new MySqlConnection(connetionString);
                cnn.Open();
                MySqlCommand command = cnn.CreateCommand();
                try
                {
                    command.CommandText = "INSERT INTO report(IDforDB,TimeStamp,Rep,CompanyName,Place,DebtorCode,DebtorName,CustomerCode,CustomerName,Speciality,Department,Contact,Rep2,BussinesUnit,Events,ReportType,Product,Objective,Outcome,POA,notesDebtor,notesCustomer,Location,Lat,Longitude,SWver,DataVer,MailDate,MailFrom)" +
                    " VALUES(?n1,?n2,?n3,?n4,?n5,?n6,?n7,?n8,?n9,?n10,?n11,?n12,?n13,?n14,?n15,?n16,?n17,?n18,?n19,?n20,?n21,?n22,?n23,?n24,?n25,?n26,?n27,?n28,?n29)";


                    command.Parameters.Add("?n1", MySqlDbType.VarChar).Value = this.textBoxID.Text;
                    
                    command.Parameters.Add("?n2", MySqlDbType.VarChar).Value = this.textBoxTimeStamp.Text;
                    command.Parameters.Add("?n3", MySqlDbType.VarChar).Value = this.textBoxRep.Text;
                    
                    command.Parameters.Add("?n4", MySqlDbType.VarChar).Value = this.textBoxCompany.Text;
                    command.Parameters.Add("?n5", MySqlDbType.VarChar).Value = this.textBoxPlace.Text;
                    command.Parameters.Add("?n6", MySqlDbType.VarChar).Value = this.textBoxDebtorCode.Text;
                    command.Parameters.Add("?n7", MySqlDbType.VarChar).Value = this.textBoxDebtorName.Text;
                    command.Parameters.Add("?n8", MySqlDbType.VarChar).Value = this.textBoxCustomerCode.Text;
                    command.Parameters.Add("?n9", MySqlDbType.VarChar).Value = this.textBoxCustomerName.Text;
                    command.Parameters.Add("?n10", MySqlDbType.VarChar).Value = this.textBoxSpeciality.Text;
                    command.Parameters.Add("?n11", MySqlDbType.VarChar).Value = this.textBoxDepartment.Text;
                    command.Parameters.Add("?n12", MySqlDbType.VarChar).Value = this.textBoxPharmacy.Text;
                    command.Parameters.Add("?n13", MySqlDbType.VarChar).Value = this.textBoxRep2.Text;
                    command.Parameters.Add("?n14", MySqlDbType.VarChar).Value = this.textBoxBussinesUnit.Text;
                    command.Parameters.Add("?n15", MySqlDbType.VarChar).Value = this.textBoxEvents.Text;
                    command.Parameters.Add("?n16", MySqlDbType.VarChar).Value = this.textBoxReportType.Text;
                    command.Parameters.Add("?n17", MySqlDbType.VarChar).Value = this.textBoxProduct.Text;
                    command.Parameters.Add("?n18", MySqlDbType.Text).Value = this.textBoxObjective.Text;
                    command.Parameters.Add("?n19", MySqlDbType.Text).Value = this.textBoxOutcome.Text;
                    command.Parameters.Add("?n20", MySqlDbType.Text).Value = this.textBoxPOA.Text;
                    command.Parameters.Add("?n21", MySqlDbType.VarChar).Value = this.textBoxnotesDebtor.Text;
                    command.Parameters.Add("?n22", MySqlDbType.VarChar).Value = this.textBoxnotesCustomer.Text;
                    command.Parameters.Add("?n23", MySqlDbType.VarChar).Value = this.textBoxLocation.Text;
                    command.Parameters.Add("?n24", MySqlDbType.VarChar).Value = this.textBoxLat.Text;
                    command.Parameters.Add("?n25", MySqlDbType.VarChar).Value = this.textBoxLong.Text;
                    command.Parameters.Add("?n26", MySqlDbType.VarChar).Value = this.textBoxSWver.Text;
                    command.Parameters.Add("?n27", MySqlDbType.VarChar).Value = this.textBoxDataVer.Text;
                    command.Parameters.Add("?n28", MySqlDbType.VarChar).Value = this.FromtextBox.Text;
                    command.Parameters.Add("?n29", MySqlDbType.VarChar).Value = this.DatetextBox.Text;
                    command.ExecuteNonQuery();
                    cnn.Close();
                    
                }
                catch (Exception ex)
                {

                    textBoxHealth.Text = sbdb.ToString();
                    MessageBox.Show(sbdb.ToString());
                }
            }
            //=======extract draft================
            void ExstarctDraft(string Body)
            {
                int fname = 0;
                int daysold = 0;
                char[] delimiterChars = { ';', '\n', '\r', '~' };
                string rdata = Body;
                string[] words = rdata.Split(delimiterChars);
                string Wcode;
                int len = words.Count();

                for (int i = 0; i < len; i++)
                {
                    Wcode = words[i].ToString();
                    switch (Wcode)
                    {
                        case "Report ":
                            fname = i + 1;
                            daysold = i + 2;
                            break;

                        default:
                            break;
                    }
                }
                textBoxDrafts.Text = words[fname].ToString();
                textBoxDaysOld.Text = words[daysold].ToString();


                WriteToExecldraft();
            }
            //=====================================================================
            void WriteToExecldraft()
            {


                sb.Clear();

                sb.Append(MailFrom);
                sb.Append(MailDate);
                sb.Append(textBoxDrafts.Text);
                sb.Append(textBoxDaysOld.Text);


                Microsoft.Office.Interop.Excel.Application oXL;
                Microsoft.Office.Interop.Excel._Workbook oWB;
                Microsoft.Office.Interop.Excel._Worksheet oSheet;
                Microsoft.Office.Interop.Excel.Range oRng;
                object misvalue = System.Reflection.Missing.Value;

                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = oXL.Workbooks.Open("c:\\data\\draftmail.xlsx");
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                oSheet.Cells[RowCount, 1] = MailFrom;
                oSheet.Cells[RowCount, 2] = MailDate;
                oSheet.Cells[RowCount, 3] = textBoxDrafts.Text;
                oSheet.Cells[RowCount, 4] = textBoxDaysOld.Text;


                //oRng.EntireColumn.AutoFit();

                oXL.Visible = false;
                oXL.UserControl = false;
                oWB.Save();

                oWB.Close();
                RowCount++;

            }
        }

       
        //===================================================================


    }
}


