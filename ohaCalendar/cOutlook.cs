//using Microsoft.Graph.Beta;
//using Microsoft.Graph.Beta.Models;
//using Microsoft.Identity.Client;
using Microsoft.Data.SqlClient;
using ohaCalendar.Properties;
using System.DirectoryServices;
using System.Globalization;
using System.Net.Mail;
using System.Reflection;
using System.Text;

namespace ohaCalendar
{
    public class cOutlook
    {


        // s. http://msdn.microsoft.com/en-us/library/aa908088.aspx
        public enum OlBusyStatus
        {
            olFree = 0,
            olTentative = 1,
            olBusy = 2,
            olOutOfOffice = 3
        }

        // s. http://msdn.microsoft.com/en-us/library/aa911624.aspx
        public enum OlImportance
        {
            olImportanceLow = 0,
            olImportanceNormal = 1,
            olImportanceHigh = 2
        }

        // s. http://msdn.microsoft.com/en-us/library/bb208072.aspx
        public enum OlFolderType
        {
            olFolderCalendar = 9 // The Calendar folder. 
    ,
            olFolderConflicts = 19 // The Conflicts folder (subfolder of Sync Issues folder). Only available for an Exchange account. 
    ,
            olFolderContacts = 10 // The Contacts folder. 
    ,
            olFolderDeletedItems = 3 // The Deleted Items folder. 
    ,
            olFolderDrafts = 16 // The Drafts folder. 
    ,
            olFolderInbox = 6 // The Inbox folder. 
    ,
            olFolderJournal = 11 // The Journal folder. 
    ,
            olFolderJunk = 23 // The Junk E-Mail folder. 
    ,
            olFolderLocalFailures = 21 // The Local Failures folder (subfolder of Sync Issues folder). Only available for an Exchange account. 
    ,
            olFolderManagedEmail = 29 // The top-level folder in the Managed Folders group. For more information on Managed Folders, see Help in Microsoft Outlook. Only available for an Exchange account. 
    ,
            olFolderNotes = 12 // The Notes folder. 
    ,
            olFolderOutbox = 4 // The Outbox folder. 
    ,
            olFolderSentMail = 5 // The Sent Mail folder. 
    ,
            olFolderServerFailures = 22 // The Server Failures folder (subfolder of Sync Issues folder). Only available for an Exchange account. 
    ,
            olFolderSyncIssues = 20 // The Sync Issues folder. Only available for an Exchange account. 
    ,
            olFolderTasks = 13 // The Tasks folder. 
    ,
            olFolderToDo = 28 // The To Do folder. 
    ,
            olPublicFoldersAllPublicFolders = 18 // The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account. 
    ,
            olFolderRssFeeds = 25 // The RSS Feeds folder. 
        }

        // s. http://www.online-excel.de/excel/singsel_vba.php?f=85
        public enum OlItemType
        {
            olMailItem = 0,
            olAppointmentItem = 1,
            olContactItem = 2,
            olTaskItem = 3,
            olJournalItem = 4,
            olNoteItem = 5,
            olPostItem = 6,
            olDistributionListItem = 7
        }

        public class OlAttachmentType
        {
            public const Int32 olByReference = 4;
            public const Int32 olByValue = 1;
            public const Int32 olEmbeddedItem = 5;
            public const Int32 olOLE = 6;
        }

        public enum OlUserPropertyType
        {
            olCombination = 19,  // The Property type Is a combination Of other types. It corresponds To the MAPI type PT_STRING8.
            olCurrency = 14, // Represents a Currency Property type. It corresponds To the MAPI type PT_CURRENCY.
            olDateTime = 5,  // Represents a DateTime Property type. It corresponds To the MAPI type PT_SYSTIME.
            olDuration = 7,  // Represents a time duration Property type. It corresponds To the MAPI type PT_LONG.
            olEnumeration = 21,  // Represents an enumeration Property type. It corresponds To the MAPI type PT_LONG.
            olFormula = 18,  // Represents a formula Property type. It corresponds To the MAPI type PT_STRING8. See UserDefinedProperty.Formula Property.
            olInteger = 20,  // Represents an Integer number Property type. It corresponds To the MAPI type PT_LONG.
            olKeywords = 11, // Represents a String array Property type used To store keywords. It corresponds To the MAPI type PT_MV_STRING8.
            olNumber = 3, // Represents a Double number Property type. It corresponds To the MAPI type PT_DOUBLE.
            olOutlookInternal = 0,   // Represents an Outlook internal Property type.
            olPercent = 12,  // Represents a Double number Property type used To store a percentage. It corresponds To the MAPI type PT_LONG.
            olSmartFrom = 22,    // Represents a smart from Property type. This Property indicates that If the From Property Of an Outlook item Is empty, Then the To Property should be used instead.
            olText = 1,  // Represents a String Property type. It corresponds To the MAPI type PT_STRING8.
            olYesNo = 6 // Represents a yes/no (Boolean) property type. It corresponds to the MAPI type PT_BOOLEAN.
        }

        const string m_c_postfach_invoice = "Postfach - Invoice";
        private string m_entryID;
        private Int32 m_AttachmentsIndex;
        private Int32? m_docu_archive_selectiontype;
        private Int32? m_serien_numberssysid;
        private Int32? m_docu_archive_docutype;
        private SqlConnection m_connection;
        private Int32 m_client;

        private dynamic m_Mails;
        private string m_Current_culture;

        private dynamic m_Application;
        private dynamic m_outlookApp;
        private dynamic m_MailItem;
        private dynamic m_TaskItem;


        private string m_subject;
        private string m_bodydescription;
        private dynamic m_AttachmentsSource;
        private bool m_is_fax;

        public struct structDocu_archive
        {
            //DateTime? m_document_date;
            //public DateTime? Document_date { set { m_document_date = value; } get { return m_document_date; } }
            int? m_docuarchiveheadsysid;
            public int? Docuarchiveheadsysid { set { m_docuarchiveheadsysid = value; } get { return m_docuarchiveheadsysid; } }
            string m_Path;
            public string Path { set { m_Path = value; } get { return m_Path; } }
            string m_docuarchivefilename;
            public string Docuarchivefilename { set { m_docuarchivefilename = value; } get { return m_docuarchivefilename; } }
            string m_description;
            public string Description { set { m_description = value; } get { return m_description; } }
            DateTime? m_docuarchivevaliduntil;
            public DateTime? Docuarchivevaliduntil { set { m_docuarchivevaliduntil = value; } get { return m_docuarchivevaliduntil; } }
            string m_docuarchive_version;
            public string Docuarchive_version { set { m_docuarchive_version = value; } get { return m_docuarchive_version; } }
            bool? m_is_copy;
            public bool? Is_copy { set { m_is_copy = value; } get { return m_is_copy; } }
            private int? m_language_documentsysid;
            public int? Language_documentsysid { set { m_language_documentsysid = value; } get { return m_language_documentsysid; } }

            public int? docuware_docid { set; get; }

            public int? Bas_docu_archive_invoicesysid { set; get; }

            // QM
            private int? m_qmsample_testreport_headersysid;
            public int? Qmsample_testreport_headersysid { set { m_qmsample_testreport_headersysid = value; } get { return m_qmsample_testreport_headersysid; } }
            private int? m_qmqualitycontrolplanheadersysid;
            public int? Qmqualitycontrolplanheadersysid { set { m_qmqualitycontrolplanheadersysid = value; } get { return m_qmqualitycontrolplanheadersysid; } }
            private int? m_qmnotice_of_defects_headersysid;
            public int? Qmnotice_of_defects_headersysid { set { m_qmnotice_of_defects_headersysid = value; } get { return m_qmnotice_of_defects_headersysid; } }


            private int? m_docuarchivedocutypesysid;
            public int? Docuarchivedocutypesysid { get { return m_docuarchivedocutypesysid; } set { m_docuarchivedocutypesysid = value; } }
            private int? m_docuarchiveseriennumberssysid;
            public int? Seriennumberssysid { get { return m_docuarchiveseriennumberssysid; } set { m_docuarchiveseriennumberssysid = value; } }
            private int? m_docuarchiveselectiontypesysid;
            public int? Docuarchiveselectiontypesysid { get { return m_docuarchiveselectiontypesysid; } set { m_docuarchiveselectiontypesysid = value; } }
            private int? m_docuarchivefiletypesysid;
            public int? Docuarchivefiletypesysid { get { return m_docuarchivefiletypesysid; } set { m_docuarchivefiletypesysid = value; } }
            private int? m_docuarchivedocumentsysid;
            public int? Docuarchivedocumentsysid { get { return m_docuarchivedocumentsysid; } set { m_docuarchivedocumentsysid = value; } }

            // PP
            private int? m_ppworkingplanheadersysid;
            public int? PPworkingplanheadersysid { set { m_ppworkingplanheadersysid = value; } get { return m_ppworkingplanheadersysid; } }

            private int? m_ppwork_placesysid;
            public int? PPwork_placesysid { set { m_ppwork_placesysid = value; } get { return m_ppwork_placesysid; } }

            // BAS
            private int? m_basarticlesysid;
            public int? Basarticlesysid { set { m_basarticlesysid = value; } get { return m_basarticlesysid; } }
            private int? m_basarticle_price_calculation_headersysid;
            public int? Basarticle_price_calculation_headersysid { set { m_basarticle_price_calculation_headersysid = value; } get { return m_basarticle_price_calculation_headersysid; } }
            private int? m_basarticledrawssysid;
            public int? Basarticledrawssysid { set { m_basarticledrawssysid = value; } get { return m_basarticledrawssysid; } }
            private int? m_basgenericdocumentsysid;
            public int? Basgenericdocumentsysid { set { m_basgenericdocumentsysid = value; } get { return m_basgenericdocumentsysid; } }
            private bool? m_basgenericdocument_active;
            public bool? Basgenericdocument_active { set { m_basgenericdocument_active = value; } get { return m_basgenericdocument_active; } }
            private int? m_bastooltrunksysid;
            public int? Bastooltrunksysid { set { m_bastooltrunksysid = value; } get { return m_bastooltrunksysid; } }
            private int? m_bastooltrunkstatesysid;
            public int? Bastooltrunkstatesysid { set { m_bastooltrunkstatesysid = value; } get { return m_bastooltrunkstatesysid; } }
            private int? m_bastooltrunkdrawssysid;
            public int? Bastooltrunkdrawssysid { set { m_bastooltrunkdrawssysid = value; } get { return m_bastooltrunkdrawssysid; } }
            private int? m_basorderheadersysid;
            public int? Basorderheadersysid { set { m_basorderheadersysid = value; } get { return m_basorderheadersysid; } }
            private int? m_basreclamation_headersysid;
            public int? Basreclamation_headersysid { set { m_basreclamation_headersysid = value; } get { return m_basreclamation_headersysid; } }
            private int? m_basitemslist_headersysid;
            public int? Basitemslist_headersysid { set { m_basitemslist_headersysid = value; } get { return m_basitemslist_headersysid; } }

            // CRM
            private int? m_crmassociationsysid;
            public int? Crmassociationsysid { set { m_crmassociationsysid = value; } get { return m_crmassociationsysid; } }
            private int? m_crmcustomersysid;
            public int? Crmcustomersysid { set { m_crmcustomersysid = value; } get { return m_crmcustomersysid; } }
            private int? m_crmproposition_headersysid;
            public int? Crmproposition_headersysid { set { m_crmproposition_headersysid = value; } get { return m_crmproposition_headersysid; } }
            private int? m_crmproposal_price_headersysid;
            public int? Crmproposal_price_headersysid { set { m_crmproposal_price_headersysid = value; } get { return m_crmproposal_price_headersysid; } }
            private int? m_crmcalculation_preversion_headersysid;
            public int? Crmcalculation_preversion_headersysid { set { m_crmcalculation_preversion_headersysid = value; } get { return m_crmcalculation_preversion_headersysid; } }
            private int? m_crmcalculationdrawsysid;
            public int? Crmcalculationdrawsysid { set { m_crmcalculationdrawsysid = value; } get { return m_crmcalculationdrawsysid; } }
            private int? m_crminquiry_headersysid;
            public int? Crminquiry_headersysid { set { m_crminquiry_headersysid = value; } get { return m_crminquiry_headersysid; } }
            private int? m_crminquiry_projectsysid;
            public int? Crminquiry_projectsysid { set { m_crminquiry_projectsysid = value; } get { return m_crminquiry_projectsysid; } }
            private int? m_crmorderheadersysid;
            public int? Crmorderheadersysid { set { m_crmorderheadersysid = value; } get { return m_crmorderheadersysid; } }
            private int? m_crmcommissionheadersysid;
            public int? Crmcommissionheadersysid { set { m_crmcommissionheadersysid = value; } get { return m_crmcommissionheadersysid; } }
            private int? m_order_header_confirmation_headersysid;
            public int? Crmorder_header_confirmation_headersysid { set { m_order_header_confirmation_headersysid = value; } get { return m_order_header_confirmation_headersysid; } }
            private int? m_crminvoice_headersysid;
            public int? Crminvoice_headersysid { set { m_crminvoice_headersysid = value; } get { return m_crminvoice_headersysid; } }
            private int? m_crmorder_details_callssysid;
            public int? Crmorder_details_callssysid { set { m_crmorder_details_callssysid = value; } get { return m_crmorder_details_callssysid; } }

            // SCM
            private int? m_scmorder_detail_arrivalsysid;
            public int? Scmorder_detail_arrivalsysid { set { m_scmorder_detail_arrivalsysid = value; } get { return m_scmorder_detail_arrivalsysid; } }
            private int? m_scmsuppliersysid;
            public int? Scmsuppliersysid { set { m_scmsuppliersysid = value; } get { return m_scmsuppliersysid; } }
            private int? m_scmorderheadersysid;
            public int? Scmorderheadersysid { set { m_scmorderheadersysid = value; } get { return m_scmorderheadersysid; } }
            public int? Scmorder_header_bill_of_deliverysysid { set { m_scmorder_header_bill_of_deliverysysid = value; } get { return m_scmorder_header_bill_of_deliverysysid; } }
            private int? m_scmreturnheadersysid;
            public int? Scmreturnheadersysid { set { m_scmreturnheadersysid = value; } get { return m_scmreturnheadersysid; } }
            private int? m_scmorder_header_bill_of_deliverysysid;


            // HR
            private int? m_hrstaffsysid;

            public int? Hrstaffsysid { set { m_hrstaffsysid = value; } get { return m_hrstaffsysid; } }

            public Dictionary<string, int?> GetValues()
            {
                Dictionary<string, int?> return_value = new Dictionary<string, int?>();
                int tmp_int = -1;
                Type structDocu_archive_Type = typeof(structDocu_archive);

                try
                {
                    PropertyInfo[] myFields = structDocu_archive_Type.GetProperties();

                    for (int i = 0; i < myFields.Length; i++)
                    {
                        string key_name = myFields[i].Name;
                        PropertyInfo oPropertyInfo = structDocu_archive_Type.GetProperty(key_name);
                        object tmp_value = oPropertyInfo.GetValue(this, null);
                        if (tmp_value != null &&
                            int.TryParse(tmp_value.ToString(), out tmp_int))
                        {
                            return_value.Add(key_name, tmp_int);
                        }
                        else
                        {
                            return_value.Add(key_name, null);
                        }
                    }
                    return return_value;
                }
                catch
                {
                    return null;
                }
            }


        }
        public structDocu_archive m_structDocu_archive = new structDocu_archive();



        public cOutlook()
        {
        }

        public cOutlook(string entryID, Int32 AttachmentsIndex, SqlConnection connection, Int32 client, dynamic Mails, string Current_culture)
        {
            m_entryID = entryID;
            m_AttachmentsIndex = AttachmentsIndex;
            m_connection = connection;
            m_client = client;
            m_Mails = Mails;
            m_Current_culture = Current_culture;
        }

        public cOutlook(string entryID, Int32 AttachmentsIndex, Int32? docu_archive_selectiontype, Int32? serien_numberssysid, Int32? docu_archive_docutype, SqlConnection connection, Int32 client, dynamic Mails, string Current_culture)
        {
            m_entryID = entryID;
            m_AttachmentsIndex = AttachmentsIndex;
            m_docu_archive_selectiontype = docu_archive_selectiontype;
            m_serien_numberssysid = serien_numberssysid;
            m_docu_archive_docutype = docu_archive_docutype;
            m_connection = connection;
            m_client = client;
            m_Mails = Mails;
            m_Current_culture = Current_culture;
        }

        private void InitByArrayList(
            string pTo,
            string Cc,
            string Bcc,
            dynamic AttachmentsSource,
            dynamic ReadReceiptRequested,
            dynamic VotingApply,
            string subject,
            string bodydescription)
        {
            try
            {
                m_MailItem.To = pTo;
                if (!string.IsNullOrEmpty(Cc))
                    m_MailItem.CC = Cc;

                if (!string.IsNullOrEmpty(Bcc))
                    m_MailItem.BCC = Bcc;

                m_MailItem.ReadReceiptRequested = ReadReceiptRequested;

                if (VotingApply)
                    m_MailItem.VotingOptions = "Accept;Reject";

                if (!string.IsNullOrEmpty(subject))
                {
                    m_MailItem.Subject = subject;

                    if (m_is_fax == false)
                        m_MailItem.Body = bodydescription;// EMAIL
                }


                if (AttachmentsSource != null)
                {
                    if (AttachmentsSource is System.Collections.ArrayList)
                    {
                        foreach (dynamic attach_obj in AttachmentsSource)
                            m_MailItem.Attachments.AddHandler(attach_obj, OlAttachmentType.olByValue, 1, attach_obj);
                    }
                    else
                        m_MailItem.Attachments.Add(AttachmentsSource, OlAttachmentType.olByValue, 1, AttachmentsSource.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public bool CheckOutlook()
        {
            dynamic objApp = null;
            dynamic objEmail;
            try
            {
                //objApp = Interaction.CreateObject("Outlook.Application", "");
                //objEmail = objApp.CreateItem(OlItemType.olMailItem);

                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                objApp = Activator.CreateInstance(OutlookType);
                objEmail = objApp.CreateItem(OlItemType.olMailItem);

                return objEmail != null;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                objEmail = null;
                objApp = null;
            }
        }

        // s. http://support.microsoft.com/kb/285202/de
        enum OlBodyFormat
        {
            olFormatUnspecified = 0,
            olFormatPlain = 1,
            olFormatHTML = 2,
            olFormatRichText = 3
        }


        public void SendMail(string server, string _to, string _from, string subject, string body, List<string> AttachmentsPath)
        {
            try
            {
                MailMessage Message = new MailMessage(_from, _to, subject, body);
                SmtpClient client = new SmtpClient(server);
                client.Port = 25;
                client.Timeout = 10000;
                Message.IsBodyHtml = true;
                Message.Bcc.Add(new MailAddress(_from));

                foreach (string Path in AttachmentsPath)
                    Message.Attachments.Add(new System.Net.Mail.Attachment(Path));

                client.Send(Message);

                Message.Attachments.Dispose();
                Message.Dispose();
            }
            catch (Exception ex)
            {
                string message = Resources.msgModul + ": Error" + Environment.NewLine + Resources.msgMethod + ": " + ex.TargetSite.ToString() + Environment.NewLine + ex.Message;
                MessageBox.Show(message, Resources.msgCaptionError, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public bool SendEmailPerOutlook(
             string pTo,
             string Subject,
             string Body,
             List<string>? AttachmentsPath,
             string Staff_EMail = "",
             string BCC = "",
             bool Display = false,
             string CC = "")
        {
            dynamic objEmailOutlook;
            dynamic objApp;

            try
            {
                Type ExcelType = Type.GetTypeFromProgID("Outlook.Application");
                objApp = Activator.CreateInstance(ExcelType);
                objEmailOutlook = objApp.CreateItem(OlItemType.olMailItem);

                objEmailOutlook.To = pTo;

                Encoding _encoding = Encoding.GetEncoding("ISO-8859-1");
                byte[] _bytes = _encoding.GetBytes(Subject);
                string uuEncoded = Convert.ToBase64String(_bytes);
                string _encoded_subject = _encoding.GetString(System.Convert.FromBase64String(uuEncoded)).Trim();

                objEmailOutlook.BodyFormat = 2; //HTML
                objEmailOutlook.Subject = _encoded_subject; // Subject

                string _signature = objEmailOutlook.HTMLBody;
                objEmailOutlook.HTMLBody = Body + "<br>" + _signature;    //'.Replace(Chr(10), "<br>");

                if (AttachmentsPath != null)
                {
                    foreach (string Path in AttachmentsPath)
                        objEmailOutlook.Attachments.Add(Path);
                }

                if (!string.IsNullOrEmpty(Staff_EMail))
                    objEmailOutlook.To = Staff_EMail;

                if (!string.IsNullOrEmpty(BCC))
                    objEmailOutlook.BCC = BCC;

                if (!string.IsNullOrEmpty(CC))
                    objEmailOutlook.CC = CC;

                if (Display)
                    objEmailOutlook.Display();
                else
                {
                    objEmailOutlook.Send();
                    MessageBox.Show("EMail wurde erfolgreich gesendet", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                objEmailOutlook = null;
                objApp = null;
            }
        }


        //public bool SendEmailPerOutlook(
        //    string pTo, string Subject,
        //    string Body,
        //    List<string>? AttachmentsPath,
        //    string Staff_EMail = "",
        //    string BCC = "",
        //    bool Display = false,
        //    string CC = "")
        //{
        //    dynamic objEmailOutlook;
        //    dynamic objApp;
        //    string _signature = "";

        //    try
        //    {
        //        Type ExcelType = Type.GetTypeFromProgID("Outlook.Application");
        //        objApp = Activator.CreateInstance(ExcelType);
        //        objEmailOutlook = objApp.CreateItem(OlItemType.olMailItem);

        //        if (Display)
        //        {
        //            objEmailOutlook.Display();
        //            _signature = objEmailOutlook.HTMLBody;
        //        }

        //        objEmailOutlook.To = pTo;

        //        Encoding _encoding = Encoding.GetEncoding("ISO-8859-1");
        //        byte[] _bytes = _encoding.GetBytes(Subject);
        //        string uuEncoded = Convert.ToBase64String(_bytes);
        //        string _encoded_subject = _encoding.GetString(System.Convert.FromBase64String(uuEncoded)).Trim();

        //        objEmailOutlook.BodyFormat = 2; //HTML
        //        objEmailOutlook.Subject = _encoded_subject; // Subject
        //        objEmailOutlook.HTMLBody = Body + "<br>" + _signature;    //'.Replace(Chr(10), "<br>");


        //        if (AttachmentsPath != null)
        //        {
        //            foreach (string Path in AttachmentsPath)
        //                objEmailOutlook.Attachments.Add(Path);
        //        }

        //        if (!string.IsNullOrEmpty(Staff_EMail))
        //            objEmailOutlook.To = Staff_EMail;

        //        if (!string.IsNullOrEmpty(BCC))
        //            objEmailOutlook.BCC = BCC;

        //        if (!string.IsNullOrEmpty(CC))
        //            objEmailOutlook.CC = CC;

        //        if (Display)
        //            objEmailOutlook.Display();
        //        else
        //        {
        //            objEmailOutlook.Send();
        //            MessageBox.Show("EMail wurde erfolgreich gesendet", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        }
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        return false;
        //    }
        //    finally
        //    {
        //        objEmailOutlook = null;
        //        objApp = null;
        //    }
        //}


        //public bool SendEmailPerOutlookByMailto(
        //    string pTo,
        //    string Subject,
        //    string Body,
        //    List<string> AttachmentsPath,
        //    string Staff_EMail = "",
        //    string BCC = "",
        //    bool Display = true,
        //    string CC = "")
        //{
        //    try
        //    {
        //        // Dim _encoding As Encoding = Encoding.GetEncoding("ISO-8859-1")
        //        // Dim _bytes As Byte() = _encoding.GetBytes(Subject)
        //        // Dim uuEncoded As String = Convert.ToBase64String(_bytes)
        //        // Dim _encoded_subject As String
        //        // _encoded_subject = _encoding.GetString(System.Convert.FromBase64String(uuEncoded)).Trim()
        //        // Dim _subject As String = Uri.EscapeUriString(_encoded_subject)
        //        // Dim _body As String = Uri.EscapeUriString(Body)

        //        Body = Body.Replace("<br>", "");

        //        var mapi = new SimpleMapi();

        //        string Attaches = null;
        //        mapi.AddRecipient(name: pTo, addr: "", cc: false);
        //        if (!string.IsNullOrEmpty(CC))
        //            mapi.AddRecipient(name: CC, addr: "", cc: true);
        //        if (!string.IsNullOrEmpty(BCC))
        //            mapi.AddRecipientBCC(name: BCC, addr: "");
        //        if (AttachmentsPath != null && AttachmentsPath.Count > 0)
        //        {
        //            foreach (string path in AttachmentsPath)
        //                Attaches += path + ";";
        //            Attaches = Attaches.Substring(0, Attaches.Length - 1);
        //            if (!string.IsNullOrEmpty(Attaches))
        //                mapi.Attach(filepath: Attaches);
        //        }
        //        mapi.Send(subject: Subject, noteText: Body, ShowDialog: Display);

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        return false;
        //    }
        //}

        private bool SetEmailSignature(ref dynamic mail)
        {
            string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\\Microsoft\\Signatures";
            string signature = string.Empty;
            DirectoryInfo diInfo = new DirectoryInfo(appDataDir);

            if ((diInfo.Exists))
            {
                try
                {
                    string[] _directories = Directory.GetDirectories(appDataDir);

                    if ((_directories.Length == 1))
                    {
                        // Get Row-Signature
                        FileInfo[] fiSignature = diInfo.GetFiles("*.htm");

                        if ((fiSignature.Length > 0))
                        {
                            StreamReader sr = new StreamReader(fiSignature[0].FullName, Encoding.Default);
                            signature = sr.ReadToEnd();

                            if ((string.IsNullOrEmpty(signature) == true))
                                return false;
                        }
                        else
                            return false;

                        try
                        {
                            if ((string.IsNullOrEmpty(mail.HTMLBody) == false))
                            {
                                string[] _image_jpg_file = Directory.GetFiles(_directories[0], "*.jpg");
                                if ((_image_jpg_file.Length == 1))
                                {
                                    if ((File.Exists(_image_jpg_file[0])))
                                    {
                                        const string SchemaPR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com//mapi//proptag//0x3712001E";
                                        dynamic attach = mail.Attachments.Add(_image_jpg_file[0]);
                                        attach.PropertyAccessor.SetProperty(SchemaPR_ATTACH_CONTENT_ID, Path.GetFileName(_image_jpg_file[0]));
                                        mail.HTMLBody = GetBodyWithoutSalutation(mail.HTMLBody) + "<br><br>" + signature;
                                        mail.HTMLBody = SetImageSRC(mail.HTMLBody, Path.GetFileName(_image_jpg_file[0]));
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            mail.HTMLBody = GetBodyWithoutSalutation(mail.HTMLBody) + "<br><br>" + signature;
                            return false;
                        }
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            return false;
        }

        private dynamic SetImageSRC(string HTMLBody, string FileName)
        {
            string _ret_val = string.Empty;
            int _src; // "SRC:"-String
            int _brace_left, _brace_right;     // Klammer
            string _old_src_content;
            int _lenght = HTMLBody.Length;

            _src = HTMLBody.IndexOf("src");

            _brace_left = HTMLBody.IndexOf(char.ConvertFromUtf32(34), _src);
            _brace_right = HTMLBody.IndexOf(char.ConvertFromUtf32(34), _brace_left + 1);

            _old_src_content = HTMLBody.Substring(_brace_left + 1, _brace_right - _brace_left - 1);
            _ret_val = HTMLBody.Replace(_old_src_content, FileName);

            return _ret_val;
        }

        private string GetBodyWithoutSalutation(string html_body)
        {
            string _tmp_str = string.Empty;

            _tmp_str = html_body.Substring(0, html_body.LastIndexOf("<br>"));
            _tmp_str = _tmp_str.Substring(0, _tmp_str.LastIndexOf("<br>"));

            return _tmp_str;
        }

        public static dynamic getFolderByName(string FolderName, dynamic ParentFolder, ref string EntryID, ref string StoreID)
        {
            dynamic Application;
            dynamic oNameSpace;
            dynamic oFolder;

            EntryID = string.Empty;
            StoreID = string.Empty;

            //Application = Interaction.CreateObject("Outlook.Application", "");
            //oNameSpace = Application.GetNamespace("MAPI");

            Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
            Application = Activator.CreateInstance(OutlookType);
            oNameSpace = Application.GetNamespace("MAPI");

            if (ParentFolder == null)
                oFolder = oNameSpace.Folders(FolderName);
            else
                oFolder = ParentFolder.Folders(FolderName);

            if (oFolder != null)
            {
                EntryID = oFolder.EntryID;
                StoreID = oFolder.StoreID;
            }

            return oFolder;
        }

        public static void Move_EMail(string EntryIDFolder
          , string EntryIDStore
          , string EntryID_Item
          , dynamic DestinyFolder_obj
          , bool ShowMessageBox)
        {
            dynamic Application;
            dynamic oNameSpace;
            dynamic oMails;
            dynamic oMailItem;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");
                //oNameSpace = Application.GetNamespace("MAPI");

                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                oNameSpace = Application.GetNamespace("MAPI");

                oMails = oNameSpace.GetFolderFromID(EntryIDFolder, EntryIDStore).Items;

                if (oMails != null)
                {
                    for (Int32 i = 1; i <= oMails.Count; i++)
                    {
                        try
                        {
                            oMailItem = oMails(i);
                        }
                        catch
                        {
                            oMailItem = DBNull.Value;
                        }


                        if (oMailItem.EntryID == EntryID_Item)
                        {
                            oMailItem.Move(DestinyFolder_obj);

                            if (ShowMessageBox)
                                MessageBox.Show(Resources.email_deleted);

                            return;
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void DeleteEMail(string EntryIDFolder
           , string EntryIDStore
           , string EntryID_Item)
        {
            dynamic Application;
            dynamic oNameSpace;
            dynamic oMails;
            dynamic oMailItem;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");
                //oNameSpace = Application.GetNamespace("MAPI");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                oNameSpace = Application.GetNamespace("MAPI");
                oMails = oNameSpace.GetFolderFromID(EntryIDFolder, EntryIDStore).Items;

                if (oMails != null)
                {
                    for (Int32 i = 1; i <= oMails.Count; i++)
                    {
                        try
                        {
                            oMailItem = oMails(i);
                        }
                        catch
                        {
                            oMailItem = DBNull.Value;
                        }


                        if (oMailItem.EntryID == EntryID_Item)
                        {
                            oMailItem.Delete();
                            // MessageBox.Show(Resources.mgDeleteEmailItem_OK)
                            return;
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    //    public static void GetEMails(emailsDataTable emailsDataTable
    //        , attachmentsDataTable attachmentsDataTable
    //        , string EntryIDFolder
    //        , string EntryIDStore
    //        , ref dynamic Mails
    //        , bool ShowAll
    //)
    //    {
    //        dynamic Application;
    //        dynamic oNameSpace;
    //        dynamic oMails;
    //        dynamic oAttachments;
    //        dynamic oMailItem;

    //        try
    //        {
    //            //Application = Interaction.CreateObject("Outlook.Application", "");
    //            Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
    //            Application = Activator.CreateInstance(OutlookType);

    //            emailsDataTable.Rows.Clear();
    //            attachmentsDataTable.Rows.Clear();

    //            oNameSpace = Application.GetNamespace("MAPI");
    //            oMails = oNameSpace.GetFolderFromID(EntryIDFolder, EntryIDStore).Items;

    //            if (oMails != null)
    //            {
    //                for (Int32 i = 1; i <= oMails.Count; i++)
    //                {
    //                    try
    //                    {
    //                        oMailItem = oMails(i);
    //                    }
    //                    catch
    //                    {
    //                        oMailItem = DBNull.Value;
    //                    }

    //                    // If Not IsDBNull(oMailItem) And oMailItem.Attachments.Count > 0 Then
    //                    emailsRow emailsRow = emailsDataTable.NewemailsRow();

    //                    try
    //                    {
    //                        emailsRow.from = oMailItem.SenderName;
    //                        emailsRow.subject = oMailItem.Subject;
    //                        emailsRow.body = oMailItem.Body;
    //                        emailsRow.recived = oMailItem.ReceivedTime;
    //                        emailsRow.id = oMailItem.EntryID;

    //                        // EMails
    //                        emailsDataTable.AddemailsRow(emailsRow);

    //                        // Attachments
    //                        oAttachments = oMailItem.Attachments;

    //                        for (Int32 j = 1; j <= oAttachments.Count; j++)
    //                        {
    //                            dynamic oAttachment = oAttachments(j);

    //                            try
    //                            {
    //                                attachmentsDataTable.AddattachmentsRow(oAttachment.DisplayName, oAttachment.FileName, oAttachment.Index, emailsRow);
    //                            }
    //                            catch (Exception ex)
    //                            {
    //                            }
    //                        }
    //                    }
    //                    catch (Exception ex)
    //                    {
    //                    }
    //                }

    //                Mails = oMails;
    //            }
    //        }

    //        catch (Exception ex)
    //        {
    //            MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    //        }
    //    }

        public void GetOutlookMessageFile(ref string FileName
                 , ref string Subject)
        {
            dynamic Application;
            dynamic oMailItem;
            string filetype = string.Empty;
            string file_path = string.Empty;
            string file_path_tmp = string.Empty;
            string Docu_archive_searchtextsysid;

            try
            {
                Docu_archive_searchtextsysid = string.Empty;

                //Application = Interaction.CreateObject("Outlook.Application", "");

                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                file_path_tmp = Path.Combine(Path.GetTempPath(), file_path);

                for (Int32 i = 1; i <= m_Mails.Count; i++)
                {
                    oMailItem = m_Mails(i);
                    if (oMailItem.EntryID == m_entryID)
                    {
                        file_path = Path.Combine(file_path_tmp, Guid.NewGuid().ToString() + ".msg");
                        FileName = file_path;
                        Subject = oMailItem.Subject;


                        bool _is_filename_ok = false;
                        string _only_file_name = Path.GetFileName(FileName);
                        string _only_file_name_path = Path.GetDirectoryName(FileName);
                        string _only_file_name_adjust = getAdjustPath(_only_file_name, ref _is_filename_ok).Trim();

                        _only_file_name_adjust = _is_filename_ok ? _only_file_name_adjust : Guid.NewGuid().ToString() + Path.GetExtension(FileName);

                        FileName = Path.Combine(_only_file_name_path, _only_file_name_adjust);

                        oMailItem.SaveAs(file_path);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string? getAdjustPath(string Input, ref bool is_ok)
        {
            var _tmp_res = System.Text.RegularExpressions.Regex.Replace(Input, @"[\\/:*?""<>|]", string.Empty);
            is_ok = Input == _tmp_res;
            return _tmp_res;
        }

        public void GetAttachmentsFile(string EntryID
                    , ref string FileName
                    , ref string EMail
                    , ref string Subject
                    , ref string Body
                    , ref dynamic Attachments)
        {
            dynamic Application;
            dynamic oMailItem;
            dynamic oAttachments;
            string filetype = string.Empty;
            string file_path = string.Empty;
            string file_path_tmp = string.Empty;
            // Dim ValidUntil As DateTime?
            string Docu_archive_searchtextsysid;
            // Dim is_copy As Boolean

            try
            {
                Docu_archive_searchtextsysid = string.Empty;

                //Application = Interaction.CreateObject("Outlook.Application", "");

                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                file_path_tmp = Path.Combine(Path.GetTempPath(), file_path);

                for (Int32 i = 1; i <= m_Mails.Count; i++)
                {
                    oMailItem = m_Mails(i);

                    try
                    {
                        if (oMailItem != null && oMailItem.EntryID != null && oMailItem.EntryID == EntryID)
                        {
                            oAttachments = oMailItem.Attachments;

                            if (oAttachments.Count > 0)
                            {
                                if ((string.IsNullOrEmpty(oMailItem.SenderEmailAddress)))
                                    EMail = oMailItem.SenderName;
                                else
                                    EMail = oMailItem.SenderEmailAddress;

                                Subject = oMailItem.Subject;
                                Body = oMailItem.Body;
                                Attachments = oMailItem.Attachments;

                                if (Attachments != null)
                                {
                                    foreach (dynamic oAttachment in oAttachments)
                                    {
                                        if (oAttachment.Index == m_AttachmentsIndex)
                                        {
                                            file_path = Path.Combine(file_path_tmp, oAttachment.FileName.ToString().Trim().Replace(" ", "_"));
                                            FileName = file_path;
                                            oAttachment.SaveAsFile(file_path);

                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        public void GetAttachmentsFile(bool g_admin
                    , dynamic ocArchive
                    , ref string FileName
                    , ref string EMail
                    , ref string Subject
                    , ref string Body
                    , ref dynamic Attachments)
        {
            dynamic Application;
            dynamic oMailItem;
            dynamic oAttachments;
            string filetype = string.Empty;
            string file_path = string.Empty;
            string file_path_tmp = string.Empty;
            // Dim ValidUntil As DateTime?
            string Docu_archive_searchtextsysid;
            // Dim is_copy As Boolean

            try
            {
                Docu_archive_searchtextsysid = string.Empty;
                ocArchive.m_structDocu_archive = m_structDocu_archive;

                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);

                file_path_tmp = Path.Combine(Path.GetTempPath(), file_path);

                for (Int32 i = 1; i <= m_Mails.Count; i++)
                {
                    oMailItem = m_Mails(i);

                    if (oMailItem != null && oMailItem.EntryID != null && oMailItem.EntryID == m_entryID)
                    {
                        oAttachments = oMailItem.Attachments;

                        if (oAttachments.Count > 0)
                        {
                            EMail = oMailItem.SenderEmailAddress;
                            Subject = oMailItem.Subject;
                            Body = oMailItem.Body;
                            Attachments = oMailItem.Attachments;

                            if (Attachments != null)
                            {
                                foreach (dynamic oAttachment in oAttachments)
                                {
                                    if (oAttachment.Index == m_AttachmentsIndex)
                                    {
                                        file_path = Path.Combine(file_path_tmp, oAttachment.FileName);
                                        FileName = file_path;
                                        oAttachment.SaveAsFile(file_path);

                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void GetAttachmentsPathAndType(ref string FilePath, ref string type)
        {
            dynamic Application;
            dynamic oMailItem;
            dynamic oAttachments;
            string filetype = string.Empty;
            string file_path = string.Empty;
            string file_path_tmp = string.Empty;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");

                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                file_path_tmp = Path.Combine(Path.GetTempPath(), file_path);

                for (Int32 i = 1; i <= m_Mails.Count; i++)
                {
                    oMailItem = m_Mails(i);

                    if (oMailItem.EntryID == m_entryID)
                    {
                        oAttachments = oMailItem.Attachments;

                        if (oAttachments.Count > 0)
                        {
                            foreach (dynamic oAttachment in oAttachments)
                            {
                                if (oAttachment.Index == m_AttachmentsIndex)
                                {
                                    file_path = Path.Combine(file_path_tmp, oAttachment.FileName);
                                    oAttachment.SaveAsFile(file_path);
                                    filetype = GetFileNameOrExtension(file_path, true);

                                    FilePath = file_path;
                                    type = filetype;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string GetFileNameOrExtension(string pFullFileName, bool IsExtension)
        {
            string Ret_Value;
            Int32 _LastSlashPos;
            Int32 _Lenght = pFullFileName.Length;

            if (IsExtension)
                _LastSlashPos = pFullFileName.LastIndexOf(".") + 1;
            else
                _LastSlashPos = pFullFileName.LastIndexOf(@"\\") + 1;

            Ret_Value = pFullFileName.Substring(_LastSlashPos, _Lenght - _LastSlashPos);

            return Ret_Value;
        }

        //private void EMAIL_BODY_GET(Int32 clientsysid, Int32 languagesysid, string document_number, SqlConnection connection, ref string subject, ref string bodydescription, bool g_admin)
        //{
        //    SqlCommand oSqlCommand = new SqlCommand();
        //    SqlDataReader oSqlDataReader = oSqlCommand.ExecuteReader();

        //    try
        //    {
        //        if (string.IsNullOrEmpty(document_number))
        //        {
        //            dbUtilities oDBUtilities = new dbUtilities(connection, "ohaBas.email_body_get");
        //            oDBUtilities.InitStoredProcedure();
        //            oDBUtilities.addParameter("@clientsysid", SqlDbType.Int, clientsysid);
        //            oDBUtilities.addParameter("@languagesysid", SqlDbType.Int, languagesysid);
        //            oDBUtilities.addParameter("@number", SqlDbType.NVarChar, document_number);
        //            oSqlDataReader = oDBUtilities.StartStoredProcedure(g_admin);

        //            if ((oSqlDataReader.HasRows))
        //            {
        //                while ((oSqlDataReader.Read()))
        //                {
        //                    if ((!IsDBNull(oSqlDataReader("subject"))))
        //                        subject = oSqlDataReader("subject").ToString();

        //                    if ((!IsDBNull(oSqlDataReader("bodydescription"))))
        //                        bodydescription = oSqlDataReader("bodydescription").ToString();
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //    }
        //    finally
        //    {
        //        if (!IsDBNull(oSqlDataReader))
        //            oSqlDataReader.Close();
        //    }
        //}

        public void AddAttachments(List<structDocu_archive> attachments_arr, string DocumentNumber, bool supplier_contacts_isfax, dynamic oDocu_archive_search_Form)
        {
            structDocu_archive ostructAddAttachment;

            // ohaERP_Office.docu_archive_search oDocu_archive_search_Form =
            // new ohaERP_Office.docu_archive_search(DocumentNumber, true, true, true, true, true);
            oDocu_archive_search_Form.AllEnabled = true;
            if ((oDocu_archive_search_Form.ShowDialog() == DialogResult.OK))
            {
                // PDF-Test (only for FAX)
                // If (supplier_contacts_isfax And oDocu_archive_search_Form.FileType <> "pdf") Then

                // //    MessageBox.Show(ohaERP_Library.Properties.Resources.ex1175_01, 
                // //        ohaERP_Library.Properties.Resources.msgCaptionWarning, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                // //else
                // //{
                // End If


                // ostructAddAttachment = m_structDocu_archive

                ostructAddAttachment = oDocu_archive_search_Form.m_StructDocu_archive;

                if ((attachments_arr.IndexOf(ostructAddAttachment) == -1))
                    attachments_arr.Add(ostructAddAttachment);
            }
        }

        //public void AddAttachments(ArrayList attachmets_arr, dynamic attachment)
        //{
        //    OpenFileDialog dlgOpenFile = new OpenFileDialog();

        //    if ((Information.IsDBNull(attachment) && dlgOpenFile.ShowDialog() == DialogResult.OK))
        //        attachmets_arr.Add(dlgOpenFile.FileName);
        //    else
        //        attachmets_arr.Add(attachment);
        //}

        public class OutlookDateType
        {
            private string m_subject;
            public string Subject
            {
                get
                {
                    return m_subject;
                }
                set
                {
                    m_subject = value;
                }
            }

            private string m_body;
            public string Body
            {
                get
                {
                    return m_body;
                }
                set
                {
                    m_body = value;
                }
            }

            private DateTime m_start;
            public DateTime Start
            {
                get
                {
                    return m_start;
                }
                set
                {
                    m_start = value;
                }
            }

            private DateTime m_end;
            public DateTime EndDate
            {
                get
                {
                    return m_end;
                }
                set
                {
                    m_end = value;
                }
            }

            private Int32 m_duration;
            public Int32 Duration
            {
                get
                {
                    return m_duration;
                }
                set
                {
                    m_duration = value;
                }
            }

            private string m_location;
            public string Location
            {
                get
                {
                    return m_location;
                }
                set
                {
                    m_location = value;
                }
            }

            private OlImportance m_importance;
            public OlImportance Importance
            {
                get
                {
                    return m_importance;
                }
                set
                {
                    m_importance = value;
                }
            }
        }



        public static dynamic GetOutlookFolder(string FolderName, OlFolderType type, string FolderNameAlternate = null)
        {
            dynamic Application = null;
            dynamic mpnNamespace = null;
            dynamic oFolders = null;
            dynamic oFolder = null;
            dynamic ReturnOblect = null;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                mpnNamespace = Application.GetNamespace("MAPI");

                //var oFolder_default = mpnNamespace.GetDefaultFolder(type);

                oFolders = mpnNamespace.Folders;

                foreach (var folder in oFolders)
                {
                    if (folder.Name.ToString().ToLower() == FolderName.ToLower())
                    {
                        oFolder = folder;
                        break;
                    }
                    if (FolderNameAlternate != null && folder.Name.ToString().ToLower() == FolderNameAlternate.ToLower())
                    {
                        oFolder = folder;
                        break;
                    }
                }

                if (oFolder == null)
                {
                    if (!FolderName.Contains("@haas.de"))
                    {

                        // FolderName += "@haas.de"

                        foreach (var folder in oFolders)
                        {
                            if (folder.Name.ToString().ToLower() == FolderName.ToLower())
                            {
                                oFolder = folder;
                                break;
                            }
                        }
                    }
                }

                if (oFolder == null)
                    throw new Exception("Kein Kalender '" + FolderName + "' wurde gefunden!");
                else
                {
                    if (type == OlFolderType.olFolderCalendar)
                    {
                        foreach (var folder in oFolder.Folders)
                        {
                            if (folder.DefaultMessageClass == "IPM.Appointment")
                            {
                                ReturnOblect = folder;
                                break;
                            }
                        }
                    }
                    else if (type == OlFolderType.olFolderContacts)
                    {
                        foreach (var folder in oFolder.Folders)
                        {
                            if (folder.DefaultMessageClass == "IPM.Contact")
                            {
                                ReturnOblect = folder;
                                break;
                            }
                        }
                    }

                    return ReturnOblect;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }
            finally
            {
                Application = null;
            }
        }

        public static List<string> GetCalendarNames()
        {
            List<string> ret_val = new List<string>();
            dynamic Application = null;
            dynamic mpnNamespace = null;
            dynamic oFolders = null;
            dynamic oFolder = null;
            dynamic ReturnOblect = null;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                mpnNamespace = Application.GetNamespace("MAPI");

                // oFolder = mpnNamespace.GetDefaultFolder(OlFolderType.olFolderContacts)

                oFolders = mpnNamespace.Folders;

                foreach (var folder in oFolders)
                {
                    if (folder.DefaultMessageClass == "IPM.Note")
                        ret_val.Add(folder.Name);
                }

                return ret_val;
            }

            // If oFolder Is Nothing Then
            // If Not FolderName.Contains("@haas.de") Then

            // 'FolderName += "@haas.de"

            // For Each folder In oFolders
            // If folder.Name.ToString().ToLower() = FolderName.ToLower() Then
            // oFolder = folder
            // Exit For
            // End If
            // Next
            // End If
            // End If

            // If oFolder Is Nothing Then
            // Throw New Exception("Kein Kalender '" && FolderName && "' wurde gefunden!")
            // Else

            // If Type = OlFolderType.olFolderCalendar Then
            // For Each folder In oFolder.Folders
            // If folder.DefaultMessageClass = "IPM.Appointment" Then
            // ReturnOblect = folder
            // Exit For
            // End If
            // Next
            // ElseIf Type = OlFolderType.olFolderContacts Then
            // For Each folder In oFolder.Folders
            // If folder.DefaultMessageClass = "IPM.Contact" Then
            // ReturnOblect = folder
            // Exit For
            // End If
            // Next
            // End If

            // Return ReturnOblect
            // End If
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }
            finally
            {
                Application = null;
            }
        }

        public static void SetTerminInOutlook(OutlookDateType OutlookDate, ref string EntryID)
        {
            dynamic Application;
            dynamic mapiNS;
            dynamic apt;

            try
            {
                // Dim Application As Microsoft.Office.Interop.Outlook.Application = New Outlook.Application
                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);

                // Dim mapiNS As Object = Application.GetNamespace("MAPI") ' Outlook.NameSpace = Application.GetNamespace("MAPI")
                mapiNS = Application.GetNamespace("MAPI");
                dynamic oCalendar;  // Outlook.MAPIFolder
                string profile = "";
                dynamic oItems;  // Outlook.Items
                bool exists = false;

                // mapiNS.Logon(profile, DBNull.Value, DBNull.Value, DBNull.Value)
                // Dim apt As Microsoft.Office.Interop.Outlook._AppointmentItem = Application.CreateItem(Outlook.OlItemType.olAppointmentItem)
                apt = Application.CreateItem(OlItemType.olAppointmentItem);

                if (string.IsNullOrEmpty(EntryID))
                {
                    apt.Subject = OutlookDate.Subject;
                    apt.Body = OutlookDate.Body;
                    apt.Start = OutlookDate.Start;
                    apt.End = OutlookDate.EndDate;
                    apt.ReminderSet = true;
                    apt.ReminderMinutesBeforeStart = 10;
                    apt.Importance = OutlookDate.Importance;
                    apt.BusyStatus = OlBusyStatus.olTentative;
                    apt.Location = OutlookDate.Location;
                    apt.Save();
                    apt.Display();
                    EntryID = apt.EntryID;
                }
                else
                {
                    oCalendar = mapiNS.GetDefaultFolder(OlFolderType.olFolderCalendar);
                    oItems = oCalendar.Items;
                    foreach (var Item in oItems)
                    {
                        if (Item.EntryID == EntryID)
                        {
                            Item.Subject = OutlookDate.Subject;
                            Item.Body = OutlookDate.Body;
                            Item.Start = OutlookDate.Start;
                            Item.End = OutlookDate.EndDate;
                            Item.ReminderSet = true;
                            Item.ReminderMinutesBeforeStart = 10;
                            Item.Importance = OutlookDate.Importance;
                            Item.BusyStatus = OlBusyStatus.olTentative;
                            Item.Location = OutlookDate.Location;
                            Item.Save();
                            Item.Display();
                            exists = true;
                            break;
                        }
                    }
                    if (exists == false)
                    {
                        if (MessageBox.Show(Resources.msg1120_01, "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.OK)
                        {
                            EntryID = string.Empty;
                            SetTerminInOutlook(OutlookDate, ref EntryID);   // New Appointment
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Application = null;
                mapiNS = null;
                apt = null;
            }
        }

        public static bool DeleteTerminInOutlook(string EntryID)
        {
            dynamic Application;
            dynamic mapiNS;
            dynamic apt;

            try
            {
                // Dim Application As Microsoft.Office.Interop.Outlook.Application = New Outlook.Application
                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);

                // Dim mapiNS As Object = Application.GetNamespace("MAPI") ' Outlook.NameSpace = Application.GetNamespace("MAPI")
                mapiNS = Application.GetNamespace("MAPI");

                // Dim oCalendar As Outlook.MAPIFolder
                dynamic oCalendar;

                string profile = "";
                dynamic oItems;
                // Dim oItems As Outlook.Items            

                oCalendar = mapiNS.GetDefaultFolder(OlFolderType.olFolderCalendar);
                oItems = oCalendar.Items;
                foreach (var Item in oItems)
                {
                    if (Item.EntryID == EntryID)
                    {
                        Item.Delete();
                        MessageBox.Show(Resources.msg1125_01, "", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);

                        return true;
                    }
                }

                MessageBox.Show(Resources.msg1130_01, "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);

                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }

            finally
            {
                Application = null;
                mapiNS = null;
                apt = null;
            }
        }

        public static dynamic GetCurrentCalendar()
        {
            dynamic mpnNamespace;
            dynamic Application;

            Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
            Application = Activator.CreateInstance(OutlookType);
            mpnNamespace = Application.GetNamespace("MAPI");

            var recipient = mpnNamespace.CurrentUser;
            dynamic sharedFolder = mpnNamespace.GetSharedDefaultFolder(recipient, OlFolderType.olFolderCalendar);
            return sharedFolder;
        }

        public static List<OutlookCalendarItemType> Outlook_GetCalendarItems(
            dynamic Calendar,
            string CalendarName,
            string SearchInSubject,
            string SearchInBody,
            DateTime Start,
            DateTime? Ende,
            bool OnlyUnRead,
            string CalendarNameAlternate = null,
            bool IsLightAlgorithmus = false
            )
        {
            dynamic Application;
            dynamic mpnNamespace;
            // Dim oCalendar As Object
            dynamic oItems;
            //dynamic oResultItems;
            List<OutlookCalendarItemType> oResultItems_obj;
            string PropTag = "https://schemas.microsoft.com/mapi/proptag/";
            // dynamic strRestriction = null;
            var date_format = CultureInfo.CreateSpecificCulture("en-US");
            OutlookCalendarItemType outlookCalendarItemType_obj;
            const string oUserPropertName = "ItemID";
            string oItemID = null;

            try
            {
                oResultItems_obj = new List<OutlookCalendarItemType>();

                if (Ende == new DateTime?() || Ende == DateTime.MinValue)
                    Ende = DateTime.Now;

                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                mpnNamespace = Application.GetNamespace("MAPI");


                if (Calendar == null && string.IsNullOrEmpty(CalendarName))
                    Calendar = mpnNamespace.GetDefaultFolder(OlFolderType.olFolderCalendar);
                //    Calendar = GetOutlookFolder(CalendarName, OlFolderType.olFolderCalendar, CalendarNameAlternate);


                if (Calendar == null)
                    throw new Exception("Kein Kalender '" + CalendarName + "' wurde gefunden!");

                oItems = Calendar.Items;
                oItems.IncludeRecurrences = true;
                oItems.Sort("[Start]", Type.Missing);
                //if (OnlyUnRead)
                //    oResultItems = oItems.Find("[UnRead] = true");
                //else
                //    oResultItems = oItems;
                //string filter = "[Start] >= '" + Start.ToString("g") + "' AND [End] <= '" + ((DateTime)Ende).ToString("g") + "'";
                string filter = "[Start] <= '" + ((DateTime)Ende).ToString("g") + "' AND [End] >= '" + Start.ToString("g") + "'";
                var restrictItems = oItems.Restrict(filter);

                foreach (var item in restrictItems)
                {
                    bool to_add = false;

                    DateTime _creationTime = Convert.ToDateTime(item.CreationTime.ToString("g").Replace("#", ""));
                    DateTime _start = Convert.ToDateTime(item.Start.ToString("g").Replace("#", ""));
                    DateTime _end = Convert.ToDateTime(item.End.ToString("g").Replace("#", ""));

                    if (!string.IsNullOrEmpty(SearchInSubject))
                    {
                        if (!item.Subject == null && item.Subject.ToString().Contains(SearchInSubject))
                            to_add = true;
                    }

                    if (!string.IsNullOrEmpty(SearchInBody) && to_add == false)
                    {
                        if (!item.Body == null && item.Body.ToString().Contains(SearchInBody))
                            to_add = true;
                    }

                    if (Start != null && Ende != null)
                    {
                        if (_start >= Start && _end <= Ende)
                            to_add = true;
                    }

                    List<string> Attachments = new List<string>();

                    try
                    {
                        if (item.Attachments.Count > 0)
                        {
                            foreach (var attachment in item.Attachments)
                            {
                                if (!attachment == null && !attachment.FileName == null && !string.IsNullOrEmpty(attachment.FileName))
                                {
                                    if (attachment.FileName.ToString().ToLower().Equals("item_info.xml"))
                                    {
                                        var new_guid = Guid.NewGuid().ToString();
                                        var file_name = Path.Combine(Path.GetTempPath(), "item_info_" + new_guid + ".xml");
                                        attachment.SaveAsFile(file_name);
                                        Attachments.Add(file_name);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                    if (to_add)
                    {
                        var itemID_prop = item.UserProperties.Find(oUserPropertName);
                        if (itemID_prop == null)
                        {
                            // Save new GUID for old item
                            item.UserProperties.Add(oUserPropertName, OlUserPropertyType.olText);
                            oItemID = Guid.NewGuid().ToString();
                            item.UserProperties["ItemID"] = oItemID;
                            item.Save();
                        }
                        else
                            oItemID = itemID_prop.Value;

                        string telephone = null;
                        // Try
                        // telephone = IIf(item.Telephone Is Nothing, Nothing, item.Telephone)
                        // Catch ex As Exception
                        // End Try

                        string Mobile_phone = null;
                        // Try
                        // Mobile_phone = IIf(item.Mobile_phone Is Nothing, Nothing, item.Mobile_phone)
                        // Catch ex As Exception
                        // End Try

                        string Email = null;
                        // Try
                        // Email = IIf(item.Email Is Nothing, Nothing, item.Email)
                        // Catch ex As Exception
                        // End Try

                        string Contact_person = null;
                        // Try
                        // Contact_person = IIf(item.Contact_person Is Nothing, Nothing, item.Contact_person)
                        // Catch ex As Exception
                        // End Try

                        string Companyname1 = null;
                        // Try
                        // Companyname1 = IIf(item.Companyname1 Is Nothing, Nothing, item.Companyname1)
                        // Catch ex As Exception
                        // End Try

                        string Street = null;
                        // Try
                        // Street = IIf(item.Street Is Nothing, Nothing, item.Street)
                        // Catch ex As Exception
                        // End Try

                        string City = null;
                        // Try
                        // City = IIf(item.City Is Nothing, Nothing, item.City)
                        // Catch ex As Exception
                        // End Try

                        string Postcode = null;
                        // Try
                        // Postcode = IIf(item.Postcode Is Nothing, Nothing, item.Postcode)
                        // Catch ex As Exception
                        // End Try

                        string Nation = null;

                        string Organizer = item.Organizer;

                        string RequiredAttendees = item.RequiredAttendees;
                        // Try
                        // Nation = IIf(item.Nation Is Nothing, Nothing, item.Nation)
                        // Catch ex As Exception
                        // End Try

                        outlookCalendarItemType_obj = new OutlookCalendarItemType(
                            item.Subject, item.Body, oItemID, _creationTime, _start, _end,
                            item.Duration, item.Location, telephone, Mobile_phone, Email,
                            Contact_person, Companyname1, Street, City, Postcode, Nation, Attachments, Organizer, RequiredAttendees, null);
                        oResultItems_obj.Add(outlookCalendarItemType_obj);
                    }
                }

                return oResultItems_obj;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                Application = null;
            }
        }

        public static void CreateAppointment(string title, string body, DateTime start, string location, bool display = false, bool allDayEvent = true)
        {
            dynamic apptItem;
            dynamic Application;
            dynamic mpnNamespace;
            dynamic oCalendar;

            //Application = Interaction.CreateObject("Outlook.Application", "");
            Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
            Application = Activator.CreateInstance(OutlookType);
            mpnNamespace = Application.GetNamespace("MAPI");
            oCalendar = mpnNamespace.GetDefaultFolder(OlFolderType.olFolderCalendar);
            apptItem = Application.CreateItem(OlItemType.olAppointmentItem);

            {
                var withBlock = apptItem;
                withBlock.Subject = title;
                withBlock.Body = body;
                withBlock.Start = DateTime.Now;
                withBlock.Location = location;
                // .End = Date.Now.AddHours(1)
                withBlock.ReminderSet = true;
                withBlock.ReminderMinutesBeforeStart = 30;
                withBlock.AllDayEvent = allDayEvent;
                withBlock.Save();
            }

            if (display)
                apptItem.Display(true);

            apptItem = null;
            Application = null;
        }

        public static string UpdateInsertAppointment(
            dynamic Calendar,
            string CalendarName,
            string ItemID,
            string title,
            string body,
            DateTime start,
            DateTime? end_,
            double? duration,
            string location,
            List<string> attachments,
            bool display = false,
            string CalendarNameAlternate = null
        )
        {
            dynamic apptItem = null;
            dynamic Application;
            dynamic mpnNamespace;
            // Dim oCalendar As Object
            dynamic oAttathments = null;
            const string oUserPropertName = "ItemID";
            string oItemID = null;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                mpnNamespace = Application.GetNamespace("MAPI");

                if (Calendar == null && string.IsNullOrEmpty(CalendarName))
                    Calendar = GetOutlookFolder(CalendarName, OlFolderType.olFolderCalendar, CalendarNameAlternate);

                if (Calendar == null)
                    throw new Exception("Kein Kalender '" + CalendarName + "' wurde gefunden!");


                if (string.IsNullOrEmpty(ItemID))
                {

                    // apptItem = Application.CreateItem(OlItemType.olAppointmentItem)
                    apptItem = Calendar.Items.Add();

                    apptItem.UserProperties.Add(oUserPropertName, OlUserPropertyType.olText);
                    oItemID = Guid.NewGuid().ToString();
                    apptItem.UserProperties["ItemID"] = oItemID;
                }
                else
                {
                    oItemID = ItemID;

                    foreach (var item in Calendar.Items)
                    {
                        var itemID_prop = item.UserProperties.Find(oUserPropertName);
                        if (itemID_prop == null)
                            continue;
                        if (itemID_prop.Value == ItemID)
                        {
                            apptItem = item;
                            break;
                        }
                    }
                }

                if (apptItem == null)
                    return null;

                // Attathments
                if (attachments.Count > 0)
                {

                    // Clear old Attathments
                    while (apptItem.Attachments.Count > 0)
                        apptItem.Attachments.Remove(1);

                    foreach (var attachment in attachments)
                        apptItem.Attachments.Add(attachment);
                }

                {
                    var withBlock = apptItem;
                    withBlock.Subject = title;
                    withBlock.Body = body;
                    withBlock.Start = start;
                    withBlock.Location = location;
                    withBlock.End = end_;
                    withBlock.Duration = duration;
                    // .ReminderSet = True
                    withBlock.ReminderMinutesBeforeStart = 30;
                    // .AllDayEvent = allDayEvent
                    withBlock.Save();
                }

                if (display)
                    apptItem.Display(true);

                return oItemID;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                apptItem = null;
                Application = null;
            }
        }

        public static bool DeleteAppointment(dynamic Calendar, string CalendarName, string ItemID, string CalendarNameAlternate = null
        )
        {
            dynamic apptItem = null;
            dynamic Application;
            dynamic mpnNamespace;
            // Dim oCalendar As Object
            dynamic oAttathments = null;
            const string oUserPropertName = "ItemID";
            string oItemID = null;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                mpnNamespace = Application.GetNamespace("MAPI");


                // oCalendar = mpnNamespace.GetDefaultFolder(OlFolderType.olFolderCalendar)

                if (Calendar == null && string.IsNullOrEmpty(CalendarName))
                    Calendar = GetOutlookFolder(CalendarName, OlFolderType.olFolderCalendar, CalendarNameAlternate);

                if (Calendar == null)
                    throw new Exception("Kein Kalender '" + CalendarName + "' wurde gefunden!");


                if (!string.IsNullOrEmpty(ItemID))
                {
                    oItemID = ItemID;

                    for (var i = Calendar.Items.Count; i >= 1; i += -1)
                    {
                        var itemID_prop = Calendar.Items(i).UserProperties.Find(oUserPropertName);
                        if (itemID_prop == null)
                            continue;
                        if (itemID_prop.Value == ItemID)
                        {
                            Calendar.Items.Remove(i);
                            return true;
                        }
                    }
                }

                if (apptItem == null)
                    return false;

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                apptItem = null;
                Application = null;
            }
        }

        public static bool DeleteAppointment(dynamic Calendar, string CalendarName, DateTime startTime, DateTime endTime, string CalendarNameAlternate = null
        )
        {
            dynamic Application;
            dynamic mpnNamespace;
            string oItemID = null;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                mpnNamespace = Application.GetNamespace("MAPI");


                // oCalendar = mpnNamespace.GetDefaultFolder(OlFolderType.olFolderCalendar)

                if (Calendar == null && string.IsNullOrEmpty(CalendarName))
                    Calendar = GetOutlookFolder(CalendarName, OlFolderType.olFolderCalendar, CalendarNameAlternate);

                if (Calendar == null)
                    throw new Exception("Kein Kalender '" + CalendarName + "' wurde gefunden!");

                var filter = "[Start] = '" + startTime.ToString("g") + "' AND [End] = '" + endTime.ToString("g") + "'";

                var CalendarItems = Calendar.Items;
                // CalendarItems.IncludeRecurrences = True
                // CalendarItems.Sort("[Start]", Type.Missing)
                var restrictItems = CalendarItems.Restrict(filter);

                if (restrictItems.Count == 0)
                    return false;

                if (restrictItems.Count == 1)
                {
                    var item = restrictItems.GetFirst();

                    item.Delete();

                    // restrictItems.Delete()

                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                // apptItem = Nothing
                Application = null;
            }
        }

        public static dynamic? GetAndOpenAppointment(string EntryID, dynamic Calendar, string CalendarName, string CalendarNameAlternate = null, bool ShowItem = true)
        {
            dynamic? Application = null;
            dynamic mpnNamespace;
            string oItemID = null;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                mpnNamespace = Application.GetNamespace("MAPI");

                if (Calendar == null && string.IsNullOrEmpty(CalendarName))
                    Calendar = mpnNamespace.GetDefaultFolder(OlFolderType.olFolderCalendar);
                else
                    Calendar = GetOutlookFolder(CalendarName, OlFolderType.olFolderCalendar, CalendarNameAlternate);

                if (Calendar == null)
                    throw new Exception("Kein Kalender '" + CalendarName + "' wurde gefunden!");

                var CalendarItems = Calendar.Items;

                for (var i = Calendar.Items.Count; i >= 1; i += -1)
                {
                    var itemID_prop = Calendar.Items(i).UserProperties.Find("ItemID");
                    if (itemID_prop == null)
                        continue;
                    if (itemID_prop.Value == EntryID)
                    {
                        if (ShowItem)
                            Calendar.Items(i).Display(true);
                        return Calendar.Items(i);
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                if (Application != null)
                    Application = null;
            }
        }

        public static List<OutlookContactsItemType> Outlook_GetContactItemsByEmail(string Email)
        {
            dynamic Application;
            dynamic mpnNamespace;
            dynamic oCalendar;
            dynamic oItems;
            List<OutlookContactsItemType> oResultItems_list;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                mpnNamespace = Application.GetNamespace("MAPI");
                oCalendar = mpnNamespace.GetDefaultFolder(OlFolderType.olFolderContacts);
                oItems = oCalendar.Items;
                oResultItems_list = new List<OutlookContactsItemType>();

                foreach (var item in oItems)
                {
                    if (item.Email1Address == Email)
                        oResultItems_list.Add(new OutlookContactsItemType(null, item.CompanyName, item.Title, item.JobTitle, item.FirstName, item.LastName, item.Birthday, item.BusinessAddressCity, item.BusinessAddressCountry, item.BusinessAddressPostalCode, item.BusinessAddressState, item.BusinessAddressStreet, item.BusinessFaxNumber, item.BusinessTelephoneNumber, item.BusinessHomePage, item.Email1Address, item.Email2Address, item.Email3Address, item.MobileTelephoneNumber, item.Body));
                }

                return oResultItems_list;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                Application = null;
            }
        }

        public static List<OutlookCalendarItemType>? Outlook_GetContactsItems()
        {
            dynamic Application;
            dynamic mpnNamespace;
            dynamic oCalendar;
            dynamic oItems;
            List<OutlookCalendarItemType> oResultItems_obj;
            OutlookCalendarItemType outlookCalendarItemType_obj;
            string oItemID = null;

            try
            {
                oResultItems_obj = new List<OutlookCalendarItemType>();

                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (OutlookType == null)
                    return null;

                Application = Activator.CreateInstance(OutlookType);
                if (Application == null)
                    return null;

                mpnNamespace = Application.GetNamespace("MAPI");
                oCalendar = mpnNamespace.GetDefaultFolder(OlFolderType.olFolderContacts);
                oItems = oCalendar.Items;

                foreach (var item in oItems)
                {
                    if (item == null)
                        continue;

                    string location = null;
                    location = Get(item, "BusinessAddress");

                    string PrimaryTelephoneNumber = null;
                    PrimaryTelephoneNumber = Get(item, "PrimaryTelephoneNumber");


                    string MobileTelephoneNumber = null;
                    MobileTelephoneNumber = Get(item, "MobileTelephoneNumber");

                    string Email1DisplayName = null;
                    Email1DisplayName = Get(item, "Email1Address");  // "Email1DisplayName");

                    string LastNameAndFirstName = null;
                    LastNameAndFirstName = Get(item, "LastNameAndFirstName");

                    string CompanyName = null;
                    CompanyName = Get(item, "CompanyName");

                    string BusinessAddressStreet = null;
                    BusinessAddressStreet = Get(item, "BusinessAddressStreet");

                    string BusinessAddressCity = null;
                    BusinessAddressCity = Get(item, "BusinessAddressCity");

                    string BusinessAddressPostalCode = null;
                    BusinessAddressPostalCode = Get(item, "BusinessAddressPostalCode");

                    string BusinessAddressCountry = null;
                    BusinessAddressCountry = Get(item, "BusinessAddressCountry");

                    string Organizer = null;
                    Organizer = Get(item, "Organizer");

                    string RequiredAttendees = null;
                    RequiredAttendees = Get(item, "RequiredAttendees");

                    if (string.IsNullOrEmpty(Email1DisplayName))
                        continue;

                    outlookCalendarItemType_obj = new OutlookCalendarItemType(
                        item.Subject,
                        item.Body,
                        oItemID,
                        default(DateTime),
                        default(DateTime),
                        default(DateTime),
                        0,
                        location,
                        PrimaryTelephoneNumber,
                        MobileTelephoneNumber,
                        Email1DisplayName,
                        LastNameAndFirstName,
                        CompanyName,
                        BusinessAddressStreet,
                        BusinessAddressCity,
                        BusinessAddressPostalCode,
                        BusinessAddressCountry,
                        null,
                        Organizer,
                        RequiredAttendees,
                        null
                        );
                    oResultItems_obj.Add(outlookCalendarItemType_obj);
                }
                return oResultItems_obj;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                Application = null;
            }
        }

        //public static async Task<List<OutlookCalendarItemType>?> Outlook_GetContactsItems_ActiveDirectory()
        //{
        //    string clientId = "9a784777-43ad-4727-8b16-08ceb56735a1";
        //    string[] scopes = { "https://graph.microsoft.com/.default" };
        //    string tenantId = "c5f5a5dc-4bcd-48b2-97c4-a108275e87ba";
        //    string clientSecret = "6.B8Q~Iw0KfHkJrqur2EYQXp7w8Ld-zCiTG9LdkT";
        //    List<OutlookCalendarItemType> oResultItems_obj = new List<OutlookCalendarItemType>();

        //    IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
        //        .Create(clientId)
        //        .WithClientSecret(clientSecret)
        //        .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
        //        .Build();

        //    var authResult = await confidentialClientApplication.AcquireTokenForClient(scopes).ExecuteAsync();
        //    var httpClient = new HttpClient();
        //    httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken);
        //    var graphClient = new GraphServiceClient(httpClient);

        //    if (graphClient == null || graphClient.Users == null)
        //        return null;

        //    try
        //    {
        //        UserCollectionResponse users = await graphClient.Users.GetAsync();

        //        if (users == null)
        //            return null;
        //        // Admin_SQL_SERVER, Adminservice, 

        //        var list = users.Value.Where(x => (!string.IsNullOrEmpty(x.Mail) && x.Mail.Contains("@")));
        //        //|| x.BusinessPhones.Count > 0 || !string.IsNullOrEmpty(x.MobilePhone)) ; 
        //        foreach (var user in list)
        //        {
        //            var _user = user.DisplayName;
        //            var _presence = (await graphClient.Users[user.Id].Presence.GetAsync());
        //            if (_presence == null)
        //                continue;
        //            var _availability = _presence.Availability;
        //            var _activity = _presence.Activity;
        //            var _statusMessage = _presence.StatusMessage;
        //            string? location = user.OfficeLocation;
        //            string? PrimaryTelephoneNumber = user.BusinessPhones.Count > 0 ? user.BusinessPhones[0] : "";
        //            string? MobileTelephoneNumber = user.MobilePhone;
        //            string? LastNameAndFirstName = user.DisplayName;
        //            string? CompanyName = "";
        //            string? BusinessAddressStreet = user.StreetAddress;
        //            string? BusinessAddressCity = user.City;
        //            string? BusinessAddressPostalCode = user.PostalCode;
        //            string BusinessAddressCountry = user.Country;

        //            var outlookCalendarItemType_obj = new OutlookCalendarItemType(
        //                null, //item.Subject,
        //                null, //item.Body,
        //                user.Id,
        //                default(DateTime),
        //                default(DateTime),
        //                default(DateTime),
        //                0,
        //                location,
        //                PrimaryTelephoneNumber,
        //                MobileTelephoneNumber,
        //                user.Mail,
        //                LastNameAndFirstName,
        //                CompanyName,
        //                BusinessAddressStreet,
        //                BusinessAddressCity,
        //                BusinessAddressPostalCode,
        //                BusinessAddressCountry,
        //                null,
        //                null, //Organizer,
        //                null, //RequiredAttendees
        //                _availability
        //                );
        //            oResultItems_obj.Add(outlookCalendarItemType_obj);

        //        }
        //        return oResultItems_obj;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //        return null;
        //    }
        //}

        //public static async Task<Presence?> Outlook_GetContact_ActiveDirectory(string email)
        //{
        //    string clientId = "9a784777-43ad-4727-8b16-08ceb56735a1";
        //    string[] scopes = { "https://graph.microsoft.com/.default" };
        //    string tenantId = "c5f5a5dc-4bcd-48b2-97c4-a108275e87ba";
        //    string clientSecret = "6.B8Q~Iw0KfHkJrqur2EYQXp7w8Ld-zCiTG9LdkT";
        //    List<OutlookCalendarItemType> oResultItems_obj = new List<OutlookCalendarItemType>();

        //    IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
        //        .Create(clientId)
        //        .WithClientSecret(clientSecret)
        //        .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
        //        .Build();

        //    var authResult = await confidentialClientApplication.AcquireTokenForClient(scopes).ExecuteAsync();
        //    var httpClient = new HttpClient();
        //    httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken);
        //    var graphClient = new GraphServiceClient(httpClient);

        //    if (graphClient == null || graphClient.Users == null)
        //        return null;

        //    try
        //    {
        //        UserCollectionResponse users = await graphClient.Users.GetAsync();

        //        if (users == null)
        //            return null;

        //        var user = users.Value.FirstOrDefault(x => x.Mail == email);
        //        if (user == null) return null;
        //        var _presence = (await graphClient.Users[user.Id].Presence.GetAsync());

        //        return _presence;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //        return null;
        //    }
        //}


        public static List<OutlookCalendarItemType>? Outlook_GetContactsItems_AddressBook()
        {
            dynamic Application;
            dynamic mpnNamespace;
            dynamic addrLists;
            dynamic oItems;
            List<OutlookCalendarItemType> oResultItems_obj = new List<OutlookCalendarItemType>();
            OutlookCalendarItemType outlookCalendarItemType_obj;
            string oItemID = null;
            dynamic folderContacts;

            try
            {
                using (DirectoryEntry rootEntry = new DirectoryEntry("LDAP://haas.de"))
                {
                    using (DirectorySearcher searcher = new DirectorySearcher(rootEntry))
                    {
                        // Legen Sie die Suchkriterien fest (hier: alle Benutzer)
                        searcher.Filter = "(objectClass=user)";
                        using (SearchResultCollection results = searcher.FindAll())
                        {
                            foreach (System.DirectoryServices.SearchResult item in results)
                            {
                                string Email1DisplayName = GetPropertiesAD(item, ADProperties.EMAILADDRESS);  // "Email1DisplayName");

                                if (string.IsNullOrEmpty(Email1DisplayName)
                                    || Email1DisplayName.StartsWith("HealthMailbox")
                                    || Email1DisplayName.StartsWith("DiscoverySearchMailbox")
                                    || Email1DisplayName.StartsWith("Migration")
                                    || Email1DisplayName.StartsWith("SystemMailbox")
                                    )
                                    continue;

                                string? location = GetPropertiesAD(item, ADProperties.MANAGER);
                                string? PrimaryTelephoneNumber = GetPropertiesAD(item, ADProperties.TELEPHONE);
                                string? MobileTelephoneNumber = GetPropertiesAD(item, ADProperties.MOBILE);
                                string? LastNameAndFirstName = GetPropertiesAD(item, ADProperties.LASTNAME) + " " + GetPropertiesAD(item, ADProperties.FIRSTNAME);
                                string? CompanyName = GetPropertiesAD(item, ADProperties.COMPANY);
                                string? BusinessAddressStreet = GetPropertiesAD(item, ADProperties.STREETADDRESS);
                                string? BusinessAddressCity = GetPropertiesAD(item, ADProperties.CITY);
                                string? BusinessAddressPostalCode = GetPropertiesAD(item, ADProperties.POSTALCODE);
                                string BusinessAddressCountry = GetPropertiesAD(item, ADProperties.COUNTRY);

                                outlookCalendarItemType_obj = new OutlookCalendarItemType(
                                    null, //item.Subject,
                                    null, //item.Body,
                                    oItemID,
                                    default(DateTime),
                                    default(DateTime),
                                    default(DateTime),
                                    0,
                                    location,
                                    PrimaryTelephoneNumber,
                                    MobileTelephoneNumber,
                                    Email1DisplayName,
                                    LastNameAndFirstName,
                                    CompanyName,
                                    BusinessAddressStreet,
                                    BusinessAddressCity,
                                    BusinessAddressPostalCode,
                                    BusinessAddressCountry,
                                    null,
                                    null, //Organizer,
                                    null, //RequiredAttendees
                                    null
                                    );
                                oResultItems_obj.Add(outlookCalendarItemType_obj);
                            }
                        }

                        return oResultItems_obj;
                    }
                }
                return oResultItems_obj;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                Application = null;
            }
        }

        private static string? GetPropertiesAD(System.DirectoryServices.SearchResult item, string property_name)
        {
            string ret_val = null;
            try
            {
                ResultPropertyValueCollection item1 = item.Properties[property_name];
                if (item1.Count == 0)
                    return null;
                ret_val = item1[0].ToString();
                return ret_val;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public T IndirectCast<T>(dynamic value, T inferTypeFrom)
        {
            return (T)value;
        }

        private static string Get(dynamic item, string PropertyName)
        {
            string ret_val = null;

            try
            {
                var itemProperties = item.ItemProperties;
                var myProp = itemProperties(PropertyName);
                ret_val = myProp == null ? null : myProp.Value;
                return ret_val;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static string UpdateInsertContacts(dynamic oFolder, OutlookContactsItemType Info)
        {
            dynamic apptItem = null;
            dynamic Application = null;
            dynamic mpnNamespace = null;
            dynamic oFolders = null;
            // Dim oFolder As Object = Nothing
            dynamic oAttathments = null;
            // Const oUserPropertName As String = "ItemID"
            string oItemID = null;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                mpnNamespace = Application.GetNamespace("MAPI");

                string oSearchContact_email = null;
                string oSearchContact_telephone = null;

                if (oFolder == null)
                {
                    MessageBox.Show("Kein Personenkonto wurde gefunden!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return null;
                }

                if (!string.IsNullOrEmpty(Info.Email1Address))
                    oSearchContact_email = oFolder.Items.Find("[Email1Address]='" + Info.Email1Address + "'");

                if (!string.IsNullOrEmpty(Info.BusinessTelephoneNumber))
                    oSearchContact_telephone = oFolder.Items.Find("[BusinessTelephoneNumber]='" + Info.BusinessTelephoneNumber + "'");

                if (!string.IsNullOrEmpty(oSearchContact_email) || !string.IsNullOrEmpty(oSearchContact_telephone))
                    return null;

                apptItem = oFolder.Items.Add();

                {
                    var withBlock = apptItem;
                    // .ItemID = Info.ItemID

                    withBlock.CompanyName = Info.CompanyName;
                    withBlock.Title = Info.Title;  // Herr
                    withBlock.JobTitle = Info.JobTitle; // SW-Entwickler
                    withBlock.FirstName = Info.FirstName;
                    withBlock.LastName = Info.LastName;
                    // .Birthday = Info.Birthday

                    withBlock.BusinessAddressCity = Info.BusinessAddressCity;
                    withBlock.BusinessAddressCountry = Info.BusinessAddressCountry;
                    withBlock.BusinessAddressPostalCode = Info.BusinessAddressPostalCode;
                    withBlock.BusinessAddressState = Info.BusinessAddressState;
                    withBlock.BusinessAddressStreet = Info.BusinessAddressStreet;

                    withBlock.BusinessFaxNumber = Info.BusinessFaxNumber;
                    withBlock.BusinessTelephoneNumber = Info.BusinessTelephoneNumber;
                    withBlock.BusinessHomePage = Info.BusinessHomePage;

                    withBlock.Email1Address = Info.Email1Address;
                    withBlock.Email2Address = Info.Email2Address;
                    withBlock.Email3Address = Info.Email3Address;

                    withBlock.MobileTelephoneNumber = Info.MobileTelephoneNumber;

                    withBlock.Body = Info.Body;

                    withBlock.Save();
                }

                return oItemID;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }
            finally
            {
                apptItem = null;
                Application = null;
            }
        }


        public static string UpdateContact(OutlookContactsItemType Info)
        {
            dynamic apptItem = null;
            dynamic Application = null;
            dynamic mpnNamespace = null;
            dynamic oFolders = null;
            dynamic oFolder = null;
            dynamic oAttathments = null;
            // Const oUserPropertName As String = "ItemID"
            string oItemID = null;

            try
            {
                //Application = Interaction.CreateObject("Outlook.Application", "");
                Type OutlookType = Type.GetTypeFromProgID("Outlook.Application");
                Application = Activator.CreateInstance(OutlookType);
                mpnNamespace = Application.GetNamespace("MAPI");

                oFolder = mpnNamespace.GetDefaultFolder(OlFolderType.olFolderContacts);

                dynamic oSearchContact_email = null;
                // Dim oSearchContact_telephone = Nothing

                if (oFolder == null)
                {
                    MessageBox.Show("Kein Personenkonto wurde gefunden!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return null;
                }

                if (!string.IsNullOrEmpty(Info.Email1Address))
                    apptItem = oFolder.Items.Find("[Email1Address]='" + Info.Email1Address + "'");

                {
                    var withBlock = apptItem;
                    // .ItemID = Info.ItemID

                    withBlock.CompanyName = Info.CompanyName;
                    withBlock.Title = Info.Title;  // Herr
                    withBlock.JobTitle = Info.JobTitle; // SW-Entwickler
                    withBlock.FirstName = Info.FirstName;
                    withBlock.LastName = Info.LastName;
                    // .Birthday = Info.Birthday

                    withBlock.BusinessAddressCity = Info.BusinessAddressCity;
                    withBlock.BusinessAddressCountry = Info.BusinessAddressCountry;
                    withBlock.BusinessAddressPostalCode = Info.BusinessAddressPostalCode;
                    withBlock.BusinessAddressState = Info.BusinessAddressState;
                    withBlock.BusinessAddressStreet = Info.BusinessAddressStreet;

                    withBlock.BusinessFaxNumber = Info.BusinessFaxNumber;
                    withBlock.BusinessTelephoneNumber = Info.BusinessTelephoneNumber;
                    withBlock.BusinessHomePage = Info.BusinessHomePage;

                    withBlock.Email1Address = Info.Email1Address;
                    withBlock.Email2Address = Info.Email2Address;
                    withBlock.Email3Address = Info.Email3Address;

                    withBlock.MobileTelephoneNumber = Info.MobileTelephoneNumber;

                    withBlock.Body = Info.Body;

                    withBlock.Save();
                }

                return oItemID;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }
            finally
            {
                apptItem = null;
                Application = null;
            }
        }
    }


    public class OutlookCalendarItemType
    {
        public string Subject;
        public string Body;
        public string EntryID;
        public DateTime CreationTime;
        public DateTime Start;
        public DateTime End_;
        public double Duration;
        public string Location;
        public string Telephone;
        public string Mobile_phone;
        public string Email;
        public string Contact_person;
        public string Companyname1;
        public string Street;
        public string City;
        public string Postcode;
        public string Nation;
        public List<string> Attachments;
        public string Organizer;
        public string RequiredAttendees;
        public string Availability;

        public OutlookCalendarItemType(string subject,
            string body,
            string entryID,
            DateTime creationTime,
            DateTime start,
            DateTime end_,
            double duration,
            string location,
            string telephone,
            string mobile_phone,
            string email,
            string contact_person,
            string companyname1,
            string street,
            string city,
            string postcode,
            string nation,
            List<string> Attachments,
            string Organizer,
            string RequiredAttendees,
            string Availability
            )
        {
            {
                var withBlock = this;
                withBlock.Subject = subject;
                withBlock.Body = body;
                withBlock.EntryID = entryID;
                withBlock.CreationTime = creationTime;
                withBlock.Start = start;
                withBlock.End_ = end_;
                withBlock.Duration = duration;
                withBlock.Location = location;
                withBlock.Telephone = telephone;
                withBlock.Mobile_phone = mobile_phone;
                withBlock.Email = email;
                withBlock.Contact_person = contact_person;
                withBlock.Companyname1 = companyname1;
                withBlock.Street = street;
                withBlock.City = city;
                withBlock.Postcode = postcode;
                withBlock.Nation = nation;
                withBlock.Attachments = Attachments;
                withBlock.Organizer = Organizer;
                withBlock.RequiredAttendees = RequiredAttendees;
                withBlock.Availability = Availability;
            }
        }

        public string DisplayMember
        {
            get
            {
                return Companyname1 + " [" + Email + "]";
            }
        }

        public string ValueMember
        {
            get
            {
                string ret_val;
                Int32 _start;
                Int32 _end;

                if (string.IsNullOrEmpty(Email))
                    return null;

                _start = Email.IndexOf("(") + 1;
                _end = Email.LastIndexOf(")");
                if ((_end > _start))
                    ret_val = Email.Substring(_start, _end - _start);
                else
                    ret_val = null;

                return ret_val;
            }
        }

        public string? Address
        {
            get
            {
                string? ret_val = null;
                ret_val = Street + ", " + Postcode + " " + City + ", " + Nation;
                ret_val = ret_val == ",  , " ? "" : ret_val;
                return ret_val;
            }
        }

        public string? Companyname1_Contact_person
        {
            get
            {
                string? ret_val = null;
                ret_val = !string.IsNullOrEmpty(Companyname1) ? Companyname1 : Contact_person;
                ret_val = ret_val == ",  , " ? "" : ret_val;
                return ret_val;
            }
        }
    }

    public class OutlookContactsItemType
    {
        public string ItemID;

        public string CompanyName;
        public string Title;  // Herr
        public string JobTitle; // SW-Entwickler
        public string FirstName;
        public string LastName;
        public DateTime Birthday;

        public string BusinessAddressCity;
        public string BusinessAddressCountry;
        public string BusinessAddressPostalCode;
        public string BusinessAddressState;
        public string BusinessAddressStreet;

        public string BusinessFaxNumber;
        public string BusinessTelephoneNumber;
        public string BusinessHomePage;

        public string? Email1Address;
        public string? Email2Address;
        public string? Email3Address;

        public string MobileTelephoneNumber;

        public string Body;


        public OutlookContactsItemType(
            string itemID,
            string companyName,
            string title,
            string jobTitle,
            string firstName,
            string lastName,
            DateTime? birthday,
            string BusinessAddressCity,
            string BusinessAddressCountry,
            string BusinessAddressPostalCode,
            string BusinessAddressState,
            string BusinessAddressStreet,
            string BusinessAddressFaxNumber,
            string BusinessAddressTelephoneNumber,
            string BusinessHomePage,
            string Email1Address,
            string Email2Address,
            string Email3Address,
            string MobileTelephoneNumber,
            string body)
        {
            this.ItemID = itemID;
            this.CompanyName = companyName;
            this.Title = title;
            this.JobTitle = jobTitle;
            this.FirstName = firstName;
            this.LastName = lastName;

            if (birthday != null)
                this.Birthday = (DateTime)birthday;

            this.BusinessAddressCity = BusinessAddressCity;
            this.BusinessAddressCountry = BusinessAddressCountry;
            this.BusinessAddressPostalCode = BusinessAddressPostalCode;
            this.BusinessAddressState = BusinessAddressState;
            this.BusinessAddressStreet = BusinessAddressStreet;

            this.BusinessFaxNumber = BusinessAddressFaxNumber;
            this.BusinessTelephoneNumber = BusinessAddressTelephoneNumber;
            this.BusinessHomePage = BusinessHomePage;

            this.Email1Address = Email1Address;
            this.Email2Address = Email2Address;
            this.Email3Address = Email3Address;

            this.MobileTelephoneNumber = MobileTelephoneNumber;

            this.Body = body;
        }
    }

    public static class ADProperties
    {
        public const String OBJECTCLASS = "objectClass";
        public const String CONTAINERNAME = "cn";
        public const String LASTNAME = "sn";
        public const String COUNTRYNOTATION = "c";
        public const String CITY = "l";
        public const String STATE = "st";
        public const String TITLE = "title";
        public const String POSTALCODE = "postalCode";
        public const String PHYSICALDELIVERYOFFICENAME = "physicalDeliveryOfficeName";
        public const String FIRSTNAME = "givenName";
        public const String MIDDLENAME = "initials";
        public const String DISTINGUISHEDNAME = "distinguishedName";
        public const String INSTANCETYPE = "instanceType";
        public const String WHENCREATED = "whenCreated";
        public const String WHENCHANGED = "whenChanged";
        public const String DISPLAYNAME = "displayName";
        public const String USNCREATED = "uSNCreated";
        public const String MEMBEROF = "memberOf";
        public const String USNCHANGED = "uSNChanged";
        public const String COUNTRY = "co";
        public const String DEPARTMENT = "department";
        public const String COMPANY = "company";
        public const String PROXYADDRESSES = "proxyAddresses";
        public const String STREETADDRESS = "streetAddress";
        public const String DIRECTREPORTS = "directReports";
        public const String NAME = "name";
        public const String OBJECTGUID = "objectGUID";
        public const String USERACCOUNTCONTROL = "userAccountControl";
        public const String BADPWDCOUNT = "badPwdCount";
        public const String CODEPAGE = "codePage";
        public const String COUNTRYCODE = "countryCode";
        public const String BADPASSWORDTIME = "badPasswordTime";
        public const String LASTLOGOFF = "lastLogoff";
        public const String LASTLOGON = "lastLogon";
        public const String PWDLASTSET = "pwdLastSet";
        public const String PRIMARYGROUPID = "primaryGroupID";
        public const String OBJECTSID = "objectSid";
        public const String ADMINCOUNT = "adminCount";
        public const String ACCOUNTEXPIRES = "accountExpires";
        public const String LOGONCOUNT = "logonCount";
        public const String LOGINNAME = "sAMAccountName";
        public const String SAMACCOUNTTYPE = "sAMAccountType";
        public const String SHOWINADDRESSBOOK = "showInAddressBook";
        public const String LEGACYEXCHANGEDN = "legacyExchangeDN";
        public const String USERPRINCIPALNAME = "userPrincipalName";
        public const String EXTENSION = "ipPhone";
        public const String SERVICEPRINCIPALNAME = "servicePrincipalName";
        public const String OBJECTCATEGORY = "objectCategory";
        public const String DSCOREPROPAGATIONDATA = "dSCorePropagationData";
        public const String LASTLOGONTIMESTAMP = "lastLogonTimestamp";
        public const String EMAILADDRESS = "mail";
        public const String MANAGER = "manager";
        public const String MOBILE = "mobile";
        public const String PAGER = "pager";
        public const String FAX = "facsimileTelephoneNumber";
        public const String HOMEPHONE = "homePhone";
        public const String MSEXCHUSERACCOUNTCONTROL = "msExchUserAccountControl";
        public const String MDBUSEDEFAULTS = "mDBUseDefaults";
        public const String MSEXCHMAILBOXSECURITYDESCRIPTOR = "msExchMailboxSecurityDescriptor";
        public const String HOMEMDB = "homeMDB";
        public const String MSEXCHPOLICIESINCLUDED = "msExchPoliciesIncluded";
        public const String HOMEMTA = "homeMTA";
        public const String MSEXCHRECIPIENTTYPEDETAILS = "msExchRecipientTypeDetails";
        public const String MAILNICKNAME = "mailNickname";
        public const String MSEXCHHOMESERVERNAME = "msExchHomeServerName";
        public const String MSEXCHVERSION = "msExchVersion";
        public const String MSEXCHRECIPIENTDISPLAYTYPE = "msExchRecipientDisplayType";
        public const String MSEXCHMAILBOXGUID = "msExchMailboxGuid";
        public const String NTSECURITYDESCRIPTOR = "nTSecurityDescriptor";
        public const String TELEPHONE = "telephonenumber";
    }

}
