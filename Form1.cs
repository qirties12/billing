using System;
using System.Collections.Generic;
using System.Data;
//using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace KphBilling
{
    public partial class Form1 : Form
    {
        DataClasses1DataContext db = new DataClasses1DataContext();

        public Form1 firstForm { get; set; }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var contractList = from ab in db.AccountBases
                               join aeb in db.AccountExtensionBases on ab.AccountId equals aeb.AccountId
                               where ab.StateCode == 0
                               && aeb.Nitec_contracttype != 1
                               select ab;

            foreach(var a in contractList)
            {
                checkedListBox1.Items.Add(a.Name);
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selected = checkedListBox1.SelectedIndex;
            if (selected != -1)
            {
                label1.Text = checkedListBox1.Items[selected].ToString();
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string contractListStr = "";
            foreach (var selectedContract in checkedListBox1.CheckedItems)
            {
                contractListStr += checkedListBox1.GetItemText(selectedContract) + "\n";
            }

            var invoiceType = comboBox1.SelectedIndex;

            if (comboBox1.SelectedIndex == 0)
            {
                IndividualInvoice();
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                MultilineInvoice();
            }
             else if (comboBox1.SelectedIndex == 2)
            {
                //do both
            }
            else
            {
                MessageBox.Show("Please select a invoice type");
            }
        }

        private void IndividualInvoice()
        {

            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            if (!Directory.Exists(desktopPath + "\\InvoiceFolder"))
            {
                Directory.CreateDirectory(desktopPath + "\\InvoiceFolder");
            }

            var invoiceFilePath = desktopPath + "\\InvoiceFolder";
            var fileName = "";
            var documentCount = 0;
            DateTime invoiceGeneratedTime = DateTime.Now;

            //Find patients who have been discharged
            #region Get Discharged Pts

            var dt = new DataTable();
            dt.Columns.Add("Contact");
            dt.Columns.Add("Account ID");
            dt.Columns.Add("Patient Name");
            dt.Columns.Add("Patient Number");
            dt.Columns.Add("Contact ID");

            var tabContract = "";
            var tabID = "";
            var tabFullName = "";
            var tabPNumber = "";
            var ptGuid = "";
            //var useGuid = "";
            var patientName = "";
            var patientDOB = "";
            var patientNumber = "";
            var patientAddress1 = "";
            var patientAddress2 = "";
            var patientAddress3 = "";
            var patientAddressCity = "";
            var patientAddressCode = "";
            string nFileName = null;
            var exRefNum = "";
            var useFile = "";
            var ptReferralDate = "";
            var ptDischargeDate = "";
            var specialityName = "";
            Guid? defaultPriceLevelID = null;


            var checkedContracts = checkedListBox1.CheckedItems;
            var checkedContractsIndex = checkedListBox1.CheckedIndices;
            List<string> checkedContractsList = new List<string>();
            List<Guid> checkedContractsGuidList = new List<Guid>();
            var dischargedAfterDate = dateTimePicker1.Value;

            foreach (var contract in checkedContracts)
            {
                checkedContractsList.Add(contract.ToString());

                var getGuid = from ab in db.AccountBases
                              where ab.Name == contract.ToString()
                              select ab.AccountId;

                foreach (var g in getGuid)
                {
                    checkedContractsGuidList.Add(g);
                }
            }

            foreach (var con in checkedContractsGuidList)
            {

                var getPatients = from ab in db.AccountBases
                                  join cb in db.ContactBases on ab.AccountId equals cb.AccountId
                                  join ceb in db.ContactExtensionBases on cb.ContactId equals ceb.ContactId
                                  where ab.AccountId == con
                                  && ceb.Nitec_NextStep == 9
                                  && ceb.Nitec_DischargeDate >= dischargedAfterDate
                                  orderby ceb.Nitec_WLINumber descending
                                  select new
                                  {
                                      tabContract = ab.Name,
                                      tabID = ab.AccountId,
                                      tabFullName = cb.FullName,
                                      tabPNumber = ceb.Nitec_WLINumber,
                                      ptGuid = cb.ContactId,
                                      dplid = ab.DefaultPriceLevelId
                                  };

                int nop = getPatients.Count();

                if (nop > 0 )
                {
                    foreach (var i in getPatients)
                    {

                        tabContract = i.tabContract;
                        tabID = i.tabID.ToString();
                        tabFullName = i.tabFullName;
                        tabPNumber = i.tabPNumber;
                        ptGuid = i.ptGuid.ToString();
                        defaultPriceLevelID = i.dplid;
                        dt.Rows.Add(tabContract, tabID, tabFullName, tabPNumber, ptGuid);

                        var patient = from cb in db.ContactBases
                                      join ceb in db.ContactExtensionBases on cb.ContactId equals ceb.ContactId
                                      where cb.ContactId == i.ptGuid
                                      select new
                                      {
                                          cb.FirstName,
                                          cb.LastName,
                                          cb.BirthDate,
                                          ceb.Nitec_WLINumber,
                                          ceb.Nitec_HospitalReferenceNumber,
                                          ceb.Nitec_ReferralDateto352,
                                          ceb.Nitec_DischargeDate
                                      };

                        var getPatientAddress = from cab in db.CustomerAddressBases
                                                where cab.ParentId == i.ptGuid
                                                && cab.AddressNumber == 1
                                                select cab;



                        foreach (var b in getPatientAddress)
                        {
                            patientAddress1 = b.Line1;
                            patientAddress2 = b.Line2;
                            patientAddress3 = b.Line3;
                            patientAddressCity = b.City;
                            patientAddressCode = b.PostalCode;
                        }

                        foreach (var a in patient)
                        {
                            patientName = a.FirstName + " " + a.LastName;
                            patientDOB = string.Format("{0:dd/MM/yyyy}", ConvertFromUTC((DateTime)a.BirthDate));
                            patientNumber = a.Nitec_WLINumber.ToString();
                            fileName = a.Nitec_WLINumber.ToString();
                            exRefNum = a.Nitec_HospitalReferenceNumber;

                            if (a.Nitec_ReferralDateto352 != null)
                            {
                                ptReferralDate = string.Format("{0:dd/MM/yyyy}", ConvertFromUTC((DateTime)a.Nitec_ReferralDateto352));
                            }
                            else
                            {
                                ptReferralDate = "Not Listed";
                            }

                            if (a.Nitec_DischargeDate != null)
                            {
                                ptDischargeDate = string.Format("{0:dd/MM/yyyy}", ConvertFromUTC((DateTime)a.Nitec_DischargeDate));
                            }
                            else
                            {
                                ptDischargeDate = "Not Listed";
                            }
                        }

                        if (exRefNum == null)
                        {
                            nFileName = fileName.Replace("/", "-");
                            useFile = nFileName;
                        }
                        else
                        {
                            useFile = exRefNum;
                        }

                        var getAppointments = from paeb in db.Nitec_patientappointmentExtensionBases
                                              join pab in db.Nitec_patientappointmentBases on paeb.Nitec_patientappointmentId equals pab.Nitec_patientappointmentId
                                              join ateb in db.Nitec_appointmenttypeExtensionBases on paeb.nitec_appointmenttypeid equals ateb.Nitec_appointmenttypeId
                                              join seb in db.Nitec_specialityExtensionBases on paeb.nitec_specialityid equals seb.Nitec_specialityId
                                              where paeb.nitec_patientid == i.ptGuid
                                              && pab.statecode == 1
                                              && pab.statuscode == 8
                                              && ateb.Nitec_name != "Investigation"
                                              && ateb.Nitec_name != "Procedure"
                                              orderby paeb.Nitec_StartDateandTime
                                              select new
                                              {
                                                  ptName = paeb.Nitec_name,
                                                  paeb.nitec_patientid,
                                                  apptName = ateb.Nitec_name,
                                                  paeb.Nitec_StartDateandTime,
                                                  specialityName = seb.Nitec_name,
                                                  aptTypeID = paeb.nitec_appointmenttypeid,
                                                  aptHrgCode = ateb.nitec_hrgcodeid

                                              };

                        var getInvestigations = from pieb in db.Nitec_patientinvestigationExtensionBases
                                                join pib in db.Nitec_patientinvestigationBases on pieb.Nitec_patientinvestigationId equals pib.Nitec_patientinvestigationId
                                                join ieb in db.Nitec_investigationExtensionBases on pieb.nitec_investigationid equals ieb.Nitec_investigationId
                                                join seb in db.Nitec_specialityExtensionBases on pieb.nitec_specialityid equals seb.Nitec_specialityId
                                                where pieb.nitec_patientid == i.ptGuid
                                                && pib.statuscode == 2
                                                orderby pieb.Nitec_InvestigationDate
                                                select new
                                                {

                                                    invName = pieb.Nitec_name,
                                                    invDate = pieb.Nitec_InvestigationDate,
                                                    invSpecialityName = seb.Nitec_name,
                                                    invHrgCode = ieb.nitec_hrgcodeid

                                                };

                        var getProcedures = from ppeb in db.Nitec_patientprocedureExtensionBases
                                            join ppb in db.Nitec_patientprocedureBases on ppeb.Nitec_patientprocedureId equals ppb.Nitec_patientprocedureId
                                            join peb in db.Nitec_procedureExtensionBases on ppeb.nitec_procedureid equals peb.Nitec_procedureId
                                            join seb in db.Nitec_specialityExtensionBases on ppeb.nitec_specialityid equals seb.Nitec_specialityId
                                            where ppeb.nitec_patientid == i.ptGuid
                                            && ppb.statuscode == 2
                                            orderby ppeb.Nitec_ProcedureDate
                                            select new
                                            {
                                                proName = ppeb.Nitec_name,
                                                proDate = ppeb.Nitec_ProcedureDate,
                                                seb.Nitec_name,
                                                proHrgCode = peb.nitec_hrgcodeid

                                            };

                        DataTable dtAppts = new DataTable();
                        dtAppts.Columns.Add("Patient");
                        var dc = new DataColumn();
                        dc.DataType = Type.GetType("System.DateTime");
                        dc.ColumnName = "Date";
                        dtAppts.Columns.Add(dc);
                        dtAppts.Columns.Add("Speciality");
                        dtAppts.Columns.Add("Description");
                        dtAppts.Columns.Add("Net Price", typeof(int));


                        foreach (var g in getAppointments)
                        {
                            decimal? aptPrice = null;
                            decimal amount;
                            //int aptCount;

                            var getAptPrice = from pplb in db.ProductPriceLevelBases
                                              where pplb.PriceLevelId == i.dplid
                                              && pplb.ProductId == g.aptHrgCode
                                              select pplb;

                            if (getAptPrice.Count() > 0)
                            {
                                foreach (var price in getAptPrice)
                                {
                                    aptPrice = price.Amount;
                                }

                                dtAppts.Rows.Add(g.ptName, g.Nitec_StartDateandTime, g.specialityName, g.apptName, aptPrice);

                                foreach (var pplb in getAptPrice)
                                {
                                    amount = (decimal)pplb.Amount_Base;
                                }
                            }
                            else
                            {
                                aptPrice = 0;
                                dtAppts.Rows.Add(g.ptName, g.Nitec_StartDateandTime, g.specialityName, g.apptName, aptPrice);
                            }
                        }

                        foreach (var g2 in getInvestigations)
                        {

                            decimal? invPrice = null;
                            decimal amount;

                            var getInvPrice = from pplb in db.ProductPriceLevelBases
                                                  where pplb.PriceLevelId == i.dplid
                                                  && pplb.ProductId == g2.invHrgCode
                                                  select pplb;

                            if (getInvPrice.Count() > 0)
                            {
                                foreach (var price in getInvPrice)
                                {

                                    invPrice = price.Amount;
                                }

                                dtAppts.Rows.Add("", g2.invDate, g2.invSpecialityName, "Investigation: " + g2.invName, invPrice);

                                foreach (var pplb in getInvPrice)
                                {
                                    amount = (decimal)pplb.Amount_Base;
                                }
                            }
                            else
                            {
                                invPrice = 0;
                                dtAppts.Rows.Add("", g2.invDate, g2.invSpecialityName, "Investigation: " + g2.invName, invPrice);

                            }

                        };

                        foreach (var g3 in getProcedures)
                        {
                            decimal? proPrice = null;
                            decimal amount;

                            var getProPrice = from pplb in db.ProductPriceLevelBases
                                              where pplb.PriceLevelId == i.dplid
                                              && pplb.ProductId == g3.proHrgCode
                                              select pplb;

                            if (getProPrice.Count() > 0)
                            {
                                foreach (var price in getProPrice)
                                {
                                    proPrice = price.Amount;
                                }

                                dtAppts.Rows.Add("", g3.proDate, g3.Nitec_name, "Procedure: " + g3.proName, proPrice);

                                foreach (var pplb in getProPrice)
                                {
                                    amount = (decimal)pplb.Amount_Base;
                                }
                            }
                            else
                            {
                                proPrice = 0;
                                dtAppts.Rows.Add("", g3.proDate, g3.Nitec_name, "Procedure: " + g3.proName, proPrice);

                            }


                        }

                        

                        if (dtAppts.Rows.Count != 0)
                        {

                            var sumObject = dtAppts.Compute("Sum([Net Price])", string.Empty); 

                            DataView view = dtAppts.AsDataView();
                            view.Sort = "Date asc";
                            DataTable sortedDT = view.ToTable();

                            Document document = new Document(PageSize.A4);

                            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(invoiceFilePath + "/" + useFile + ".pdf", FileMode.Create));

                            document.Open();

                            #region Add Logo to PDF

                            var logo = Image.GetInstance("../../Content/logo-grp.png");
                            logo.SetAbsolutePosition(480, 750);
                            logo.ScaleAbsoluteHeight(70);
                            logo.ScaleAbsoluteWidth(70);
                            document.Add(logo);

                            #endregion

                            #region Add Info to PDF

                            Paragraph invDetails = new Paragraph();
                            invDetails.SpacingBefore = 50f;
                            invDetails.SpacingAfter = 25f;

                            PdfPTable invDetailsTable = new PdfPTable(2);
                            invDetailsTable.WidthPercentage = 100;

                            PdfPCell inDateCell = new PdfPCell(new Phrase("Invoice Date: " + DateTime.Now.ToShortDateString())) { HorizontalAlignment = Element.ALIGN_RIGHT, Border = Rectangle.NO_BORDER };
                            PdfPCell inNumCell = new PdfPCell(new Phrase("Invoice Number: " + useFile)) { HorizontalAlignment = Element.ALIGN_RIGHT, Border = Rectangle.NO_BORDER };
                            PdfPCell blankCell = new PdfPCell(new Phrase("")) { Border = Rectangle.NO_BORDER };

                            invDetailsTable.AddCell(blankCell);
                            invDetailsTable.AddCell(inDateCell);
                            invDetailsTable.AddCell(blankCell);
                            invDetailsTable.AddCell(inNumCell);
                            invDetails.Add(invDetailsTable);
                            document.Add(invDetails);

                            #endregion

                            #region Add Details to PDF

                            Paragraph ptDetails = new Paragraph();
                            ptDetails.SpacingBefore = 25f;
                            ptDetails.SpacingAfter = 25f;

                            PdfPTable ptDetailsTable = new PdfPTable(3);
                            ptDetailsTable.WidthPercentage = 100;

                            PdfPCell ptNameCell = new PdfPCell(new Phrase(patientName)) { HorizontalAlignment = Element.ALIGN_LEFT, Border = Rectangle.NO_BORDER };
                            PdfPCell ptDOBCell = new PdfPCell(new Phrase(patientDOB)) { HorizontalAlignment = Element.ALIGN_LEFT, Border = Rectangle.NO_BORDER };
                            PdfPCell ptAddress1Cell = new PdfPCell(new Phrase(patientAddress1)) { HorizontalAlignment = Element.ALIGN_LEFT, Border = Rectangle.NO_BORDER };
                            PdfPCell ptAddress2Cell = new PdfPCell(new Phrase(patientAddress2)) { HorizontalAlignment = Element.ALIGN_LEFT, Border = Rectangle.NO_BORDER };
                            PdfPCell ptAddress3Cell = new PdfPCell(new Phrase(patientAddress3)) { HorizontalAlignment = Element.ALIGN_LEFT, Border = Rectangle.NO_BORDER };
                            PdfPCell ptAddressCityCell = new PdfPCell(new Phrase(patientAddressCity)) { HorizontalAlignment = Element.ALIGN_LEFT, Border = Rectangle.NO_BORDER };
                            PdfPCell ptAddressCodeCell = new PdfPCell(new Phrase(patientAddressCode)) { HorizontalAlignment = Element.ALIGN_LEFT, Border = Rectangle.NO_BORDER };

                            ptDetailsTable.AddCell(ptNameCell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(ptDOBCell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(ptAddress1Cell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(ptAddress2Cell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(ptAddress3Cell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(ptAddressCityCell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(ptAddressCodeCell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetailsTable.AddCell(blankCell);
                            ptDetails.Add(ptDetailsTable);
                            document.Add(ptDetails);

                            #endregion

                            #region Add Ref Details to PDF

                            Paragraph refDetails = new Paragraph();
                            refDetails.SpacingBefore = 25f;
                            refDetails.SpacingAfter = 25f;

                            PdfPTable refDetailsTable = new PdfPTable(2);
                            refDetailsTable.WidthPercentage = 100;

                            PdfPCell yourRefCell = new PdfPCell(new Phrase("Your Ref: *****")) { HorizontalAlignment = Element.ALIGN_LEFT, Border = Rectangle.NO_BORDER };
                            PdfPCell refDateCell = new PdfPCell(new Phrase("Referral Date: " + ptReferralDate)) { HorizontalAlignment = Element.ALIGN_LEFT, Border = Rectangle.NO_BORDER };
                            PdfPCell dischargeDateCell = new PdfPCell(new Phrase("Discharge Date: " + ptDischargeDate)) { HorizontalAlignment = Element.ALIGN_LEFT, Border = Rectangle.NO_BORDER };

                            refDetailsTable.AddCell(yourRefCell);
                            refDetailsTable.AddCell(blankCell);
                            refDetailsTable.AddCell(refDateCell);
                            refDetailsTable.AddCell(blankCell);
                            refDetailsTable.AddCell(dischargeDateCell);
                            refDetailsTable.AddCell(blankCell);
                            refDetails.Add(refDetailsTable);
                            document.Add(refDetails);


                            #endregion

                            #region Add Appointments Table to PDF

                            Paragraph appointments = new Paragraph();
                            appointments.SpacingBefore = 25f;
                            appointments.SpacingAfter = 25f;

                            PdfPTable appointmentsTable = new PdfPTable(4);
                            appointmentsTable.WidthPercentage = 100;

                            PdfPCell headerCell = new PdfPCell(new Phrase("Patient Appointments")) { Colspan = 4, HorizontalAlignment = Element.ALIGN_CENTER, BackgroundColor = BaseColor.YELLOW };
                            PdfPCell totalCell = new PdfPCell(new Phrase("Total: " + sumObject)) { Colspan = 4, HorizontalAlignment = Element.ALIGN_RIGHT };
                            PdfPCell subHead1 = new PdfPCell(new Phrase("Activity Date")) { BackgroundColor = BaseColor.LIGHT_GRAY };
                            PdfPCell subHead2 = new PdfPCell(new Phrase("Speciality")) { BackgroundColor = BaseColor.LIGHT_GRAY };
                            PdfPCell subHead3 = new PdfPCell(new Phrase("Description")) { BackgroundColor = BaseColor.LIGHT_GRAY };
                            PdfPCell subHead4 = new PdfPCell(new Phrase("Net Price (£)")) { BackgroundColor = BaseColor.LIGHT_GRAY };

                            appointmentsTable.AddCell(headerCell);
                            appointmentsTable.AddCell(subHead1);
                            appointmentsTable.AddCell(subHead2);
                            appointmentsTable.AddCell(subHead3);
                            appointmentsTable.AddCell(subHead4);

                            foreach (DataRow row in sortedDT.Rows)
                            {
                                if (sortedDT.Rows.Count > 0)
                                {
                                    appointmentsTable.AddCell(new Phrase(row[1].ToString()));
                                    appointmentsTable.AddCell(new Phrase(row[2].ToString()));
                                    appointmentsTable.AddCell(new Phrase(row[3].ToString()));
                                    appointmentsTable.AddCell(new Phrase(row[4].ToString()));
                                }


                            }

                            appointmentsTable.AddCell(totalCell);


                            appointments.Add(appointmentsTable);
                            document.Add(appointments);

                            #endregion

                            document.Close();
                            document.Dispose();

                            try
                            {
                                using (PdfReader reader = new PdfReader(invoiceFilePath + "/" + useFile + ".pdf"))
                                {
                                    using (FileStream fs = new FileStream(invoiceFilePath + "/" + useFile + "_inv.pdf", FileMode.Create, FileAccess.Write, FileShare.None))
                                    {
                                        using (PdfStamper stamper = new PdfStamper(reader, fs))
                                        {
                                            int pageCount = reader.NumberOfPages;
                                            DateTime PrintTime = DateTime.Now;
                                            for (int n = 1; n <= pageCount; n++)
                                            {
                                                Rectangle pageSize = document.PageSize;
                                                var x = pageSize.GetLeft(40);
                                                var y = pageSize.GetBottom(30);
                                                var z = pageSize.GetRight(90);

                                                ColumnText.ShowTextAligned(stamper.GetOverContent(n), Element.ALIGN_CENTER, new Phrase(String.Format("Page {0} of {1}", n, pageCount)), x, y, 0);
                                                ColumnText.ShowTextAligned(stamper.GetOverContent(n), Element.ALIGN_CENTER, new Phrase(String.Format("Printed On " + PrintTime.ToString())), z, y, 0);
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                MessageBox.Show("File is already open");
                            }

                            //MessageBox.Show("Patients name is: " + patientName + ". \nPatient Number: " + patientNumber + ". \nThey have attended " + getAppointments.Count() + " appointments.\nHaving " + getInvestigations.Count() + " investigations and " + getProcedures.Count() + " procedure(s).");

                            Process.Start(invoiceFilePath + "/" + useFile + "_inv.pdf");
                            File.Delete(invoiceFilePath + "/" + useFile + ".pdf");
                            documentCount++;

                        }
                    }
                }
                else
                {
                    MessageBox.Show("No Invoices");
                }






            }

            #endregion


        }


        private void MultilineInvoice()
        {
            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            if (!Directory.Exists(desktopPath + "\\InvoiceFolder"))
            {
                Directory.CreateDirectory(desktopPath + "\\InvoiceFolder");
            }

            var invoiceFilePath = desktopPath + "\\InvoiceFolder\\";
            var fileName = "";
            var checkedContracts = checkedListBox1.CheckedItems;
            var checkedContractsIndex = checkedListBox1.CheckedIndices;
            List<string> checkedContractsList = new List<string>();
            List<Guid> checkedContractsGuidList = new List<Guid>();
            var dischargedAfterDate = dateTimePicker1.Value;
            List<string> tempFilenameList = new List<string>();

            var contractName = "";
            int? trustInt;
            string trustName = "";
            string contractSpeciality = "";
            var patientName = "";
            var ptForename = "";
            var ptSurname = "";
            var patientDOB = "";
            var patientNumber = "";
            var patientAddress1 = "";
            var patientAddress2 = "";
            var patientAddress3 = "";
            var patientAddressCity = "";
            var patientAddressCode = "";
            var exRefNum = "";
            var ptReferralDate = "";
            var ptDischargeDate = "";
            var hncNumber = "";

            var dt = new DataTable();
            dt.Columns.Add("Surname");
            dt.Columns.Add("Forename");
            //dt.Columns.Add("Type");
            dt.Columns.Add("WLI Number");
            dt.Columns.Add("Hospital No.");
            dt.Columns.Add("H&C Number");
            dt.Columns.Add("Activity Date");
            dt.Columns.Add("Speciality");
            dt.Columns.Add("Item Description");
            dt.Columns.Add("Detailed Description");

            DataSet ds = new DataSet();

            License licenseCells = new License();
            licenseCells.SetLicense("Aspose.Total.lic");
            License asposeLI = licenseCells;

            foreach (var contract in checkedContracts)
            {
                checkedContractsList.Add(contract.ToString());

                var getGuid = from ab in db.AccountBases
                              where ab.Name == contract.ToString()
                              select ab.AccountId;

                foreach (var g in getGuid)
                {
                    checkedContractsGuidList.Add(g);
                }

                //fileName = contract.ToString();
            }

            foreach (var con in checkedContractsGuidList)
            {
                var getPatients = from ab in db.AccountBases
                                  join aeb in db.AccountExtensionBases on ab.AccountId equals aeb.AccountId
                                  join cb in db.ContactBases on ab.AccountId equals cb.AccountId
                                  join ceb in db.ContactExtensionBases on cb.ContactId equals ceb.ContactId
                                  join seb in db.Nitec_specialityExtensionBases on aeb.nitec_specialityid equals seb.Nitec_specialityId
                                  where ab.AccountId == con
                                  && ceb.Nitec_NextStep == 9
                                  && ceb.Nitec_DischargeDate >= dischargedAfterDate
                                  orderby ceb.Nitec_WLINumber descending
                                  select new
                                  {
                                      tabContract = ab.Name,
                                      tabID = ab.AccountId,
                                      tabFullName = cb.FullName,
                                      tabPNumber = ceb.Nitec_WLINumber,
                                      ptGuid = cb.ContactId,
                                      dplid = ab.DefaultPriceLevelId,
                                      trustCustomer = aeb.Nitec_Trust,
                                      speciality = seb.Nitec_name,
                                      
                                  };

                int nop = getPatients.Count();

                if (nop > 0)
                {
                    foreach (var p in getPatients)
                    {
                        contractName = p.tabContract;
                        contractSpeciality = p.speciality;
                        trustInt = p.trustCustomer;


                        switch (trustInt)
                        {
                            case 1:
                                trustName = "Belfast Trust";
                                break;
                            case 8:
                                trustName = "Benenden";
                                break;
                            case 7:
                                trustName = "DVA";
                                break;
                            case 10:
                                trustName = "Health Service Executive";
                                break;
                            case 9:
                                trustName = "North West Independent Hospital";
                                break;
                            case 2:
                                trustName = "Northern Trust";
                                break;
                            case 11:
                                trustName = "Nuffield Health";
                                break;
                            case 14:
                                trustName = "PDFORRA";
                                break;
                            case 6:
                                trustName = "Private Patients";
                                break;
                            case 12:
                                trustName = "Royal College of Surgeons Ireland";
                                break;
                            case 13:
                                trustName = "Smart Care Doc";
                                break;
                            case 3:
                                trustName = "South East Trust";
                                break;
                            case 4:
                                trustName = "Southern Trust";
                                break;
                            case 5:
                                trustName = "Western Trust";
                                break;
                            default:
                                trustName = "Not Listed";
                                break;
                        }


                        var patient = from cb in db.ContactBases
                                      join ceb in db.ContactExtensionBases on cb.ContactId equals ceb.ContactId
                                      where cb.ContactId == p.ptGuid
                                      select new
                                      {
                                          cb.FirstName,
                                          cb.LastName,
                                          cb.BirthDate,
                                          ceb.Nitec_WLINumber,
                                          ceb.Nitec_HospitalReferenceNumber,
                                          ceb.Nitec_ReferralDateto352,
                                          ceb.Nitec_DischargeDate,
                                          ceb.Nitec_HealthandCareNumber
                                      };

                        var getPatientAddress = from cab in db.CustomerAddressBases
                                                where cab.ParentId == p.ptGuid
                                                && cab.AddressNumber == 1
                                                select cab;

                        foreach (var b in getPatientAddress)
                        {
                            patientAddress1 = b.Line1;
                            patientAddress2 = b.Line2;
                            patientAddress3 = b.Line3;
                            patientAddressCity = b.City;
                            patientAddressCode = b.PostalCode;
                        }

                        foreach (var a in patient)
                        {
                            patientName = a.FirstName + " " + a.LastName;
                            ptForename = a.FirstName;
                            ptSurname = a.LastName;
                            patientDOB = string.Format("{0:dd/MM/yyyy}", ConvertFromUTC((DateTime)a.BirthDate));
                            patientNumber = a.Nitec_WLINumber.ToString();
                            fileName = a.Nitec_WLINumber.ToString();
                            exRefNum = a.Nitec_HospitalReferenceNumber;
                            hncNumber = a.Nitec_HealthandCareNumber;

                            if (a.Nitec_ReferralDateto352 != null)
                            {
                                ptReferralDate = string.Format("{0:dd/MM/yyyy}", ConvertFromUTC((DateTime)a.Nitec_ReferralDateto352));
                            }
                            else
                            {
                                ptReferralDate = "Not Listed";
                            }

                            if (a.Nitec_DischargeDate != null)
                            {
                                ptDischargeDate = string.Format("{0:dd/MM/yyyy}", ConvertFromUTC((DateTime)a.Nitec_DischargeDate));
                            }
                            else
                            {
                                ptDischargeDate = "Not Listed";
                            }
                        }

                        var getAppointments = from paeb in db.Nitec_patientappointmentExtensionBases
                                              join pab in db.Nitec_patientappointmentBases on paeb.Nitec_patientappointmentId equals pab.Nitec_patientappointmentId
                                              join ateb in db.Nitec_appointmenttypeExtensionBases on paeb.nitec_appointmenttypeid equals ateb.Nitec_appointmenttypeId
                                              join seb in db.Nitec_specialityExtensionBases on paeb.nitec_specialityid equals seb.Nitec_specialityId
                                              where paeb.nitec_patientid == p.ptGuid
                                              && pab.statecode == 1
                                              && pab.statuscode == 8
                                              && ateb.Nitec_name != "Investigation"
                                              && ateb.Nitec_name != "Procedure"
                                              orderby paeb.Nitec_StartDateandTime
                                              select new
                                              {
                                                  ptName = paeb.Nitec_name,
                                                  
                                                  paeb.nitec_patientid,
                                                  apptName = ateb.Nitec_name,
                                                  paeb.Nitec_StartDateandTime,
                                                  specialityName = seb.Nitec_name,
                                                  aptTypeID = paeb.nitec_appointmenttypeid,
                                                  aptHrgCode = ateb.nitec_hrgcodeid

                                              };

                        var getInvestigations = from pieb in db.Nitec_patientinvestigationExtensionBases
                                                join pib in db.Nitec_patientinvestigationBases on pieb.Nitec_patientinvestigationId equals pib.Nitec_patientinvestigationId
                                                join ieb in db.Nitec_investigationExtensionBases on pieb.nitec_investigationid equals ieb.Nitec_investigationId
                                                join seb in db.Nitec_specialityExtensionBases on pieb.nitec_specialityid equals seb.Nitec_specialityId
                                                where pieb.nitec_patientid == p.ptGuid
                                                && pib.statuscode == 2
                                                orderby pieb.Nitec_InvestigationDate
                                                select new
                                                {

                                                    invName = pieb.Nitec_name,
                                                    invDate = pieb.Nitec_InvestigationDate,
                                                    invSpecialityName = seb.Nitec_name,
                                                    invHrgCode = ieb.nitec_hrgcodeid

                                                };

                        var getProcedures = from ppeb in db.Nitec_patientprocedureExtensionBases
                                            join ppb in db.Nitec_patientprocedureBases on ppeb.Nitec_patientprocedureId equals ppb.Nitec_patientprocedureId
                                            join peb in db.Nitec_procedureExtensionBases on ppeb.nitec_procedureid equals peb.Nitec_procedureId
                                            join seb in db.Nitec_specialityExtensionBases on ppeb.nitec_specialityid equals seb.Nitec_specialityId
                                            where ppeb.nitec_patientid == p.ptGuid
                                            && ppb.statuscode == 2
                                            orderby ppeb.Nitec_ProcedureDate
                                            select new
                                            {
                                                proName = ppeb.Nitec_name,
                                                proDate = ppeb.Nitec_ProcedureDate,
                                                seb.Nitec_name,
                                                proHrgCode = peb.nitec_hrgcodeid

                                            };

                        

                        foreach (var g in getAppointments)
                        {
                            //get columns

                            decimal? aptPrice = null;
                            decimal amount;

                            var getAptPrice = from pplb in db.ProductPriceLevelBases
                                              where pplb.PriceLevelId == p.dplid
                                              && pplb.ProductId == g.aptHrgCode
                                              select pplb;

                            if (getAptPrice.Count() > 0)
                            {
                                foreach (var price in getAptPrice)
                                {
                                    aptPrice = price.Amount;
                                }

                                dt.Rows.Add(ptSurname, ptForename, patientNumber, exRefNum, hncNumber, (string.Format("{0:dd/MM/yyyy}", ConvertFromUTC((DateTime)g.Nitec_StartDateandTime))), g.specialityName, g.apptName );

                                foreach (var pplb in getAptPrice)
                                {
                                    amount = (decimal)pplb.Amount_Base;
                                }
                            }
                            else
                            {
                                aptPrice = 0;
                                dt.Rows.Add(ptSurname, ptForename, patientNumber, exRefNum, hncNumber, (string.Format("{0:dd/MM/yyyy}", ConvertFromUTC((DateTime)g.Nitec_StartDateandTime))), g.specialityName, g.apptName );
                            }

                            var dv = new DataView(dt);
                            //dv.Sort = "Activity Date";
                            dt = dv.ToTable();
                            int aptNumber = dt.Rows.Count;

                        }

                        
                    }
                }

                fileName = contractName;
                int fn = 1;
                Workbook wb = new Workbook();

                WorksheetCollection worksheets = wb.Worksheets;
                wb.Worksheets[0].Name = fileName;

                #region Cover page


                string imgPath = "../../Content/logo-grp.png";
                byte[] imgBytes = File.ReadAllBytes(imgPath);

                MemoryStream ms = new MemoryStream();
                ms.Write(imgBytes, 0, imgBytes.Length);
                wb.Worksheets[0].Pictures.Add(0, 6, 10, 9, ms);

                #region Styles

                StyleFlag styleFlag1 = new StyleFlag();
                styleFlag1.All = true;

                Style bannerStyle = new Style();
                bannerStyle.ForegroundColor = System.Drawing.Color.FromArgb(64, 164, 213);
                bannerStyle.Pattern = BackgroundType.Solid;
                bannerStyle.VerticalAlignment = TextAlignmentType.Center;
                bannerStyle.HorizontalAlignment = TextAlignmentType.Center;

                Style underBoldStyle = new Style();
                underBoldStyle.Font.IsBold = true;
                underBoldStyle.Font.Underline = FontUnderlineType.Single;

                Style boldStyle = new Style();
                boldStyle.Font.IsBold = true;

                Style dateStyle = new Style();
                dateStyle.Custom = "dd/mm/yyyy";
                dateStyle.ShrinkToFit = true;
                //dateStyle.IsDateTime = true;
                //dateStyle.Custom()

                #endregion

                #region Ranges

                Range range1 = wb.Worksheets[0].Cells.CreateRange("A18", "I19");
                range1.Merge();
                range1.ApplyStyle(bannerStyle, styleFlag1);
                range1.PutValue(fileName.ToString(), true, true);

                #endregion

                #region Cells

                Cell cellG12 = wb.Worksheets[0].Cells["G12"];
                cellG12.PutValue("Document:");

                Cell cellI12 = wb.Worksheets[0].Cells["I12"];
                cellI12.PutValue("INVOICE");

                Cell cellG13 = wb.Worksheets[0].Cells["G13"];
                cellG13.PutValue("Document Number: ");

                Cell cellI13 = wb.Worksheets[0].Cells["I13"];
                cellI13.PutValue("1000");

                Cell cellG14 = wb.Worksheets[0].Cells["G14"];
                cellG14.PutValue("Invoice Date: ");

                Cell cellI14 = wb.Worksheets[0].Cells["I14"];
                cellI14.PutValue("31/01/2019");

                Cell cellG15 = wb.Worksheets[0].Cells["G15"];
                cellG15.PutValue("Credit Terms: ");

                Cell cellI15 = wb.Worksheets[0].Cells["I15"];
                cellI15.PutValue("30 days net");

                Cell cellA22 = wb.Worksheets[0].Cells["A22"];
                cellA22.PutValue("Invoice Details");
                cellA22.SetStyle(underBoldStyle);

                Cell cellA24 = wb.Worksheets[0].Cells["A24"];
                cellA24.PutValue("Trust: ");
                cellA24.SetStyle(boldStyle);

                Cell cellC24 = wb.Worksheets[0].Cells["C24"];
                cellC24.PutValue(trustName);

                Cell cellA25 = wb.Worksheets[0].Cells["A25"];
                cellA25.PutValue("Speciality: ");
                cellA25.SetStyle(boldStyle);

                Cell cellC25 = wb.Worksheets[0].Cells["C25"];
                cellC25.PutValue(contractSpeciality);

                Cell cellA26 = wb.Worksheets[0].Cells["A26"];
                cellA26.PutValue("Activity up to: ");
                cellA26.SetStyle(boldStyle);

                Cell cellC26 = wb.Worksheets[0].Cells["C26"];
                cellC26.PutValue(dischargedAfterDate);
                cellC26.SetStyle(dateStyle);
                

                Cell cellD26 = wb.Worksheets[0].Cells["D26"];
                cellD26.SetStyle(dateStyle);

                Cell cellA27 = wb.Worksheets[0].Cells["A27"];
                cellA27.PutValue("Invoice ID: ");
                cellA27.SetStyle(boldStyle);

                Cell cellA28 = wb.Worksheets[0].Cells["A28"];
                cellA28.PutValue("Contract Name: ");
                cellA28.SetStyle(boldStyle);

                Cell cellC28 = wb.Worksheets[0].Cells["C28"];
                cellC28.PutValue(fileName);

                Cell cellA30 = wb.Worksheets[0].Cells["A30"];
                cellA30.PutValue("Total charges for period as per details attached");

                Cell cellG32 = wb.Worksheets[0].Cells["G32"];
                cellG32.PutValue("Total Amount Due: ");
                cellG32.SetStyle(boldStyle);

                Cell cellA34 = wb.Worksheets[0].Cells["A34"];
                cellA34.PutValue("Please make payment by BACS to the following account: ");

                Cell cellA36 = wb.Worksheets[0].Cells["A36"];
                cellA36.PutValue("Bank: ");

                Cell cellA37 = wb.Worksheets[0].Cells["A37"];
                cellA37.PutValue("Account Name: ");

                Cell cellA38 = wb.Worksheets[0].Cells["A38"];
                cellA38.PutValue("Sort Code: ");

                Cell cellA39 = wb.Worksheets[0].Cells["A39"];
                cellA39.PutValue("Account Number: ");

                Cell cellC36 = wb.Worksheets[0].Cells["C36"];
                cellC36.PutValue("Barclays Bank, Donegal Square North, Belfast");

                Cell cellC37 = wb.Worksheets[0].Cells["C37"];
                cellC37.PutValue("352 Medical, LTD");

                #endregion

                #endregion

                #region Invoice page

                #region Contract Info

                Worksheet invoiceSheet = worksheets.Add("Invoice Activity Report");

                Style titleStyle = new Style();
                titleStyle.VerticalAlignment = TextAlignmentType.Center;
                titleStyle.HorizontalAlignment = TextAlignmentType.Center;
                titleStyle.Font.IsBold = true;

                Range rangeA1E2 = invoiceSheet.Cells.CreateRange("A1", "E2");
                rangeA1E2.Merge();
                rangeA1E2.ApplyStyle(titleStyle, styleFlag1);
                rangeA1E2.PutValue("Invoiced Activity Report", true, true);

                Cell cellA3 = invoiceSheet.Cells["A3"];
                cellA3.PutValue("Trust: ");

                Cell cellC3 = invoiceSheet.Cells["C3"];
                cellC3.PutValue(trustName);

                Cell cellA4 = invoiceSheet.Cells["A4"];
                cellA4.PutValue("Speciality: ");

                Cell cellC4 = invoiceSheet.Cells["C4"];
                cellC4.PutValue(contractSpeciality);

                Cell cellA5 = invoiceSheet.Cells["A5"];
                cellA5.PutValue("Activity up to: ");

                Cell cellC5 = invoiceSheet.Cells["C5"];
                cellC5.PutValue(dischargedAfterDate);
                cellC5.SetStyle(dateStyle);

                Cell cellA6 = invoiceSheet.Cells["A6"];
                cellA6.PutValue("Invoice ID: ");

                Cell cellA7 = invoiceSheet.Cells["A7"];
                cellA7.PutValue("Name: ");

                Cell cellC7 = invoiceSheet.Cells["C7"];
                cellC7.PutValue(fileName);

                #endregion


                #region Patient Table Info



                int dtRowCount = dt.Rows.Count;

                ListObjectCollection listObjects = wb.Worksheets[1].ListObjects;

                listObjects.Add("A10", "J" + dtRowCount, true);

                int tableRow = 11;

                foreach(DataRow r in dt.Rows)
                {
                    var t0 = r[0].ToString();
                    var t1 = r[1].ToString();
                    var t2 = r[2].ToString();
                    var t3 = r[3].ToString();
                    var t4 = r[4].ToString();
                    var t5 = r[5].ToString();
                    var t6 = r[6].ToString();
                    var t7 = r[7].ToString();
                    

                    invoiceSheet.Cells["A" + tableRow.ToString()].PutValue(t0);
                    invoiceSheet.Cells["B" + tableRow.ToString()].PutValue(t1);
                    invoiceSheet.Cells["C" + tableRow.ToString()].PutValue(t2);
                    invoiceSheet.Cells["D" + tableRow.ToString()].PutValue(t3);
                    invoiceSheet.Cells["E" + tableRow.ToString()].PutValue(t4);
                    invoiceSheet.Cells["F" + tableRow.ToString()].PutValue(t5);
                    invoiceSheet.Cells["H" + tableRow.ToString()].PutValue(t6);
                    invoiceSheet.Cells["I" + tableRow.ToString()].PutValue(t7);
                    
                    tableRow++;
                }

                #endregion



                #endregion



                foreach (var v in getPatients)
                {
                    //Worksheet invoiceSheet = worksheets.Add(fileName + " " + fn);
                    //fn++;


                }
                
                
                wb.Save(invoiceFilePath + fileName + ".xls", SaveFormat.Excel97To2003);
                Process.Start(invoiceFilePath + fileName + ".xls");

                
            }

            
        }


        public DateTime ConvertFromUTC(DateTime dateVar)
        {
            TimeZoneInfo timeZone = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");
            DateTime actualDT = TimeZoneInfo.ConvertTimeFromUtc(dateVar, timeZone);
            return actualDT;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
