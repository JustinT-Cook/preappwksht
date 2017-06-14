using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Collections;
using System.Collections.Specialized;

using iTextSharp.text;
using iTextSharp.text.pdf;

using EllieMae.Encompass.BusinessObjects.Loans;
using EllieMae.Encompass.Automation;
using System.Data;

namespace PreAppGen
{
    class CreatePreApp
    {

        #region Document Variables
        // The writer for the PDF
        PdfWriter writer;
        // Document we are writing to
        Document document;
        // File Stream
        FileStream fileStream;
        // Path to our document
        string filePath;
        #endregion

        #region Fonts to be used throughout document
        Font point9;
        Font point9Bold;
        Font point10;
        Font point10Bold;
        Font point10BoldWhite;
        Font point11;
        Font point11Bold;
        Font point12;
        Font point12White;
        #endregion  

        BaseColor headerBackground;

        float tableSpacingAfter = 10f;
        float estFundPadding = 5f;

        bool includePage2 = false;
        DataTable page2Table;

        // Current Loan
        Loan loan;

        // Empty Cell used for Formatting
        PdfPCell emptyCell;

        internal void CreatePDF ()
        {
            // The loan we are grabbing data form
            loan = EncompassApplication.CurrentLoan;

            // Create file at specified path
            filePath = @"C:\temp\";
            // Create path if it doesn't already exist
            DirectoryInfo di = Directory.CreateDirectory(filePath);
            // Create file
            fileStream = new FileStream(filePath + "\\" + "PreApplication.pdf", FileMode.Create);

            // Create instance of the document which represents the PDF document itself
            document = new Document(PageSize.LETTER, 20, 20, 14, 14);
            // Create instance to PDF file by creating instance of PDFWriter class
            // using the document and the filestream in the constructor
            writer = PdfWriter.GetInstance(document, fileStream);

            // Create Fonts
            point9 = FontFactory.GetFont("Calibri", 9);
            point9Bold = FontFactory.GetFont("Calibri", 9, Font.BOLD);
            point10 = FontFactory.GetFont("Calibri", 10);
            point10Bold = FontFactory.GetFont("Calibri", 10, Font.BOLD);
            point10BoldWhite = FontFactory.GetFont("Calibri", 10, Font.BOLD, BaseColor.WHITE);
            point11 = FontFactory.GetFont("Calibri", 11);
            point11Bold = FontFactory.GetFont("Calibri", 11, Font.BOLD);
            point12 = FontFactory.GetFont("Calibri", 12);
            point12White = FontFactory.GetFont("Calibri", 12, BaseColor.WHITE);

            // Create Logo Color
            headerBackground = new BaseColor(0, 220, 157);

            emptyCell = new PdfPCell(new Paragraph(" "));

            // Open document to begin writing to it
            document.Open();

            // Assemble different pieces of the PDF
            BuildPDF();

            // Release resources
            document.Close();
            writer.Close();
            fileStream.Close();

            // Open the document so that it can be saved elsewhere
            System.Diagnostics.Process.Start(filePath + "\\" + "PreApplication.pdf");
        }

        /// <summary>
        /// Calls different functions in order to properly
        /// design the document
        /// </summary>
        private void BuildPDF ()
        {
            AssembleHeader();
            BorrowerLenderData();
            CreateLoanOverview();
            Charges();
            EstFundsNeededToClose();

            if (includePage2 == true)
            {
                AdditionalCharges();
            }
        }

        /// <summary>
        /// Creates Header of document using 
        /// Image Resource
        /// </summary>
        private void AssembleHeader ()
        {
            PdfPTable table = new PdfPTable(3)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] { 2f, 3f, 5.5f };
            table.SetWidths(columnWidths);

            // LOGO
            // Get logo from resources and create an iTextSharp Image
            System.Drawing.Bitmap logo = Properties.Resources.Logo; // This must be defined in Resources

            Image imageLogo = Image.GetInstance(logo, System.Drawing.Imaging.ImageFormat.Bmp);
            // Create cell and add logo as Element.
            PdfPCell logoCell = new PdfPCell()
            {
                Border = Rectangle.NO_BORDER,
            };
            logoCell.AddElement(imageLogo);
            table.AddCell(logoCell);

            // EMPTY CELL FOR FORMATTING
            emptyCell = new PdfPCell(new Paragraph(" "))
            {
                Border = Rectangle.NO_BORDER,
            };
            emptyCell.Border = Rectangle.NO_BORDER;
            table.AddCell(emptyCell);

            // DISCLAIMER COMMENT
            PdfPCell disc = new PdfPCell(
                new Paragraph("Your actual rate, payment, and costs could be higher.\n" +
                              "Get an official Loan Estimate before choosing a loan.", point12White))
            {
                BackgroundColor = BaseColor.BLACK,
                VerticalAlignment = Element.ALIGN_MIDDLE,
            };
            table.AddCell(disc);

            table.SpacingAfter = 10f;

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to document
            document.Add(table);
        }

        #region Borrower and Lender Data

        /// <summary>
        /// Creates Table to hold information
        /// about borrow and lender
        /// </summary>
        private void BorrowerLenderData ()
        {
            PdfPTable table = new PdfPTable(3)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] { 48f, 4f, 48f};
            table.SetWidths(columnWidths);            

            // Insert Borrower Data
            table.AddCell(BorrowerData());
            
            // Empty cell for formatting
            emptyCell.Border = Rectangle.NO_BORDER;
            table.AddCell(emptyCell);

            // Prepared By Data
            table.AddCell(PreparedBy());

            table.SpacingAfter = tableSpacingAfter;

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to document
            document.Add(table);
        }

        /// <summary>
        /// Cell containing Borrower Data
        /// </summary>
        /// <returns></returns>
        private PdfPCell BorrowerData ()
        {
            PdfPTable table = new PdfPTable(2)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] { 1f, 2f};
            table.SetWidths(columnWidths);

            // HEADER
            PdfPCell header = new PdfPCell(new Paragraph("Borrower(s)", point10BoldWhite))
            {
                BackgroundColor = headerBackground,
                Border = Rectangle.BOTTOM_BORDER,
                HorizontalAlignment = 1,
                Colspan = 2
            };
            table.AddCell(header);

            // DATE
            // Label
            PdfPCell dateLab = new PdfPCell(new Paragraph("Date:", point10))
            {
                Border = Rectangle.NO_BORDER
            };
            table.AddCell(dateLab);
            // Value
            PdfPCell dateVal = new PdfPCell(new Paragraph(DateTime.Now.ToString("MM/dd/yyyy"), point10))
            {
                Border = Rectangle.NO_BORDER
            };
            table.AddCell(dateVal);

            // BORROWER(S)
            // Label
            PdfPCell borrLab = new PdfPCell(new Paragraph("Borrowers:", point10))
            {
                Border = Rectangle.NO_BORDER
            };
            table.AddCell(borrLab);
            // Value
            Paragraph borrPara = new Paragraph
            {
                new Chunk(loan.Fields["4000"].ToString() + " " + loan.Fields["4002"].ToString(), point10)
            };
            // Add Co-Borrower if they exist
            if (!string.IsNullOrEmpty(loan.Fields["4004"].ToString()))
            {
                borrPara.Add(
                    new Chunk("\n" + loan.Fields["4004"].ToString() + " " + loan.Fields["4006"].ToString(), point10));
            }
            PdfPCell borrVal = new PdfPCell(borrPara)
            {
                Border = Rectangle.NO_BORDER
            };
            table.AddCell(borrVal);

            // PROPERTY ADDRESS
            // Label
            PdfPCell propLab = new PdfPCell(new Paragraph("Property:", point10))
            {
                Border = Rectangle.NO_BORDER
            };
            table.AddCell(propLab);
            // Value
            Paragraph propPara = new Paragraph
            {
                new Chunk(loan.Fields["11"].ToString() + "\n", point10),
                new Chunk(loan.Fields["12"].ToString() + " " +
                                   loan.Fields["14"].ToString() + " " +
                                   loan.Fields["15"].ToString(), point10)
            };
            PdfPCell propVal = new PdfPCell(propPara)
            {
                Border = Rectangle.NO_BORDER
            };
            table.AddCell(propVal);

            // LOAN NUMBER
            // Label
            PdfPCell loanLab = new PdfPCell(new Paragraph("Loan #:", point10))
            {
                Border = Rectangle.NO_BORDER
            };
            table.AddCell(loanLab);
            // Value
            PdfPCell loanVal = new PdfPCell(new Paragraph(GetStringValue("364"), point10))           
            {
                Border = Rectangle.NO_BORDER
            };
            table.AddCell(loanVal);


            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to cell and remove borders
            PdfPCell cellTable = new PdfPCell(table);

            return cellTable;
        }     

        /// <summary>
        /// Cell containing Lender Data
        /// </summary>
        /// <returns></returns>
        private PdfPCell PreparedBy ()
        {
            PdfPTable table = new PdfPTable(2)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] { 1f, 4f};
            table.SetWidths(columnWidths);

            // HEADER
            PdfPCell header = new PdfPCell(new Paragraph("Prepared By", point10BoldWhite))
            {
                BackgroundColor = headerBackground,
                Border = Rectangle.BOTTOM_BORDER,
                HorizontalAlignment = 1,
                Colspan = 2
            };
            table.AddCell(header);

            // LENDER
            // Label
            PdfPCell lLabel = new PdfPCell(new Paragraph("Lender:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(lLabel);
            // Value
            Paragraph lPara = new Paragraph
            {
                new Chunk(GetStringValue("1264") + "\n", point10),
                new Chunk(GetStringValue("1257") + " " +  // Address
                GetStringValue("1258") + " " +                  // City
                GetStringValue("1259") + " " +                  // State
                GetStringValue("1260"), point10)          // Zip
            };
            PdfPCell lValue = new PdfPCell(lPara)
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(lValue);

            // LOAN ORIGINATOR
            // Label
            PdfPCell loLabel = new PdfPCell(new Paragraph("Originator:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(loLabel);
            // Value
            Paragraph loPara = new Paragraph
            {
                new Chunk(GetStringValue("1612") + "   ", point10),   // LO name
                new Chunk("NMLS: ", point10),                               // NMLS Label
                new Chunk(GetStringValue("3238"), point10)             // NMLS ID Value
            };
            PdfPCell loValue = new PdfPCell(loPara)
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(loValue);

            // PHONE
            PdfPCell pLabel = new PdfPCell(new Paragraph("Phone:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(pLabel);
            // Value
            PdfPCell pValue = new PdfPCell(new Paragraph(GetStringValue("1823"), point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(pValue);

            // EMAIL
            // Label
            PdfPCell eLabel = new PdfPCell(new Paragraph("Email:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(eLabel);
            // Value
            BaseFont bf = BaseFont.CreateFont();
            Font font = new Font(bf, 10f);
            PdfPCell eValue = new PdfPCell(new Paragraph(GetStringValue("3968"), point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(eValue);

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to cell and remove borders
            PdfPCell cellTable = new PdfPCell(table);

            return cellTable;
        }

        #endregion

        #region Loan Data 

        /// <summary>
        /// Creates Table to hold information overview of the loan
        /// </summary>
        private void CreateLoanOverview ()
        {
            PdfPTable table = new PdfPTable(3)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] { 48f, 4f, 48f};
            table.SetWidths(columnWidths);

            // Monthly Payment Data
            table.AddCell(MonthlyPayment());


            emptyCell = new PdfPCell(new Paragraph(" ", point10))
            {
                Border = Rectangle.NO_BORDER
            };
            table.AddCell(emptyCell);

            // Loan Info
            table.AddCell(LoanInfo());

            table.SpacingAfter = tableSpacingAfter;

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to document
            document.Add(table);      
        }

        /// <summary>
        /// Cell containing Monthly Payment data
        /// </summary>
        /// <returns></returns>
        private PdfPCell MonthlyPayment ()
        {
            PdfPTable table = new PdfPTable(2)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] {80f, 20f};
            table.SetWidths(columnWidths);

            // HEADER
            PdfPCell header = new PdfPCell(new Paragraph("Estimated Monthly Payment", point10BoldWhite))
            {
                BackgroundColor = headerBackground,
                Border = Rectangle.BOTTOM_BORDER,
                HorizontalAlignment = 1,
                Colspan = 4
            };
            table.AddCell(header);

            // PRINCIPAL INTEREST
            // Label
            PdfPCell piLabel = new PdfPCell(new Paragraph("Principal & Interest:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(piLabel);
            // Value
            PdfPCell piValue = new PdfPCell(new Paragraph(GetMonetaryValue("228"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(piValue);

            // OTHER/SECONDARY FINANCING (P & I)
            // Label
            PdfPCell osfLabel = new PdfPCell(new Paragraph("Other/Secondary Financing (P&I):", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(osfLabel);
            // Value
            PdfPCell osfValue = new PdfPCell(new Paragraph(GetMonetaryValue("229"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(osfValue);

            // HAZARD INSURANCE
            // Label
            PdfPCell hiLabel = new PdfPCell(new Paragraph("Hazard Insurance:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(hiLabel);
            // Value
            PdfPCell hiValue = new PdfPCell(new Paragraph(GetMonetaryValue("230"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(hiValue);

            // REAL ESTATE TAXES
            // Label
            PdfPCell retLabel = new PdfPCell(new Paragraph("Real Estate Taxes:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(retLabel);
            // Value
            PdfPCell retValue = new PdfPCell(new Paragraph(GetMonetaryValue("1405"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(retValue);

            // MORTGAGE INSURANCE
            // Label
            PdfPCell miLabel = new PdfPCell(new Paragraph("Mortgage Insurance:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(miLabel);
            // Value
            PdfPCell miValue = new PdfPCell(new Paragraph(GetMonetaryValue("232"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(miValue);

            // OTHER
            // Label
            PdfPCell oLabel = new PdfPCell(new Paragraph("Other:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(oLabel);
            // Value
            PdfPCell oValue = new PdfPCell(new Paragraph(GetMonetaryValue("234"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(oValue);

            // HOA DUES
            // Label
            PdfPCell hoaLabel = new PdfPCell(new Paragraph("HOA Dues:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(hoaLabel);
            // Value
            PdfPCell hoaValue = new PdfPCell(new Paragraph(GetMonetaryValue("233"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(hoaValue);

            // TOTAL MONTHLY PAYMENT
            // Label
            PdfPCell tmpLabel = new PdfPCell(new Paragraph("Total Monthly Payment:", point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(tmpLabel);
            // Value
            PdfPCell tmpValue = new PdfPCell(new Paragraph(GetMonetaryValue("912"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2,
                Padding = 0
            };
            table.AddCell(tmpValue);

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to cell and remove borders
            PdfPCell cellTable = new PdfPCell(table);

            return cellTable;
        }

        /// <summary>
        /// Table containing additional loan information.
        /// Returned as cell for formatting purposes
        /// </summary>
        /// <returns></returns>
        private PdfPCell LoanInfo ()
        {
            PdfPTable table = new PdfPTable(2)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] { 2f, 3f };
            table.SetWidths(columnWidths);

            // HEADER
            PdfPCell header = new PdfPCell(new Paragraph("Loan Information", point10BoldWhite))
            {
                BackgroundColor = headerBackground,
                Border = Rectangle.BOTTOM_BORDER,
                HorizontalAlignment = 1,
                Colspan = 4
            };
            table.AddCell(header);

            // LOAN TYPE
            // Label
            PdfPCell tLabel = new PdfPCell(new Paragraph("Loan Type:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(tLabel);
            // Value
            PdfPCell tValue = new PdfPCell(new Paragraph(GetStringValue("1172"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(tValue);

            // PRODUCT
            PdfPCell pLabel = new PdfPCell(new Paragraph("Product:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(pLabel);
            // Value
            PdfPCell pValue = new PdfPCell(new Paragraph(GetStringValue("LE1.X5"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(pValue);

            // Add caps if ARM
            if (GetStringValue("608") == "AdjustableRate")
            {
                // CAPS
                PdfPCell cLabel = new PdfPCell(new Paragraph("Caps:", point10))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = 0
                };
                table.AddCell(cLabel);
                // Value
                PdfPCell cValue = new PdfPCell(new Paragraph(GetPercentageValue("697") + " / " +
                                                             GetPercentageValue("695") + " / " +
                                                             GetPercentageValue("247"), point10Bold))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = 2
                };
                table.AddCell(cValue);
            }

            // TERM
            // Label
            PdfPCell termLabel = new PdfPCell(new Paragraph("Term:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(termLabel);
            // Value
            PdfPCell termValue = new PdfPCell(new Paragraph(loan.Fields["4"].ToString() + " months", point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(termValue);

            // ESTIMATED VALUE
            // Label
            PdfPCell evLabel = new PdfPCell(new Paragraph("Estimated Value:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(evLabel);
            // Value
            PdfPCell evValue = new PdfPCell(new Paragraph(NoDecimalMonetaryValue("1821"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(evValue);

            // PURCHASE PRICE
            if (GetStringValue("19") == "Purchase")
            {
                // Label
                PdfPCell ppLabel = new PdfPCell(new Paragraph("Purchase Price:", point10))
                {
                    Border = Rectangle.LEFT_BORDER,
                    HorizontalAlignment = 0
                };
                table.AddCell(ppLabel);
                // Value
                PdfPCell ppValue = new PdfPCell(new Paragraph(NoDecimalMonetaryValue("136"), point10Bold))
                {
                    Border = Rectangle.RIGHT_BORDER,
                    HorizontalAlignment = 2
                };
                table.AddCell(ppValue);
            }

            // LOAN AMOUNT
            // Label
            PdfPCell laLabel = new PdfPCell(new Paragraph("Loan Amount:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(laLabel);
            // Value
            PdfPCell laValue = new PdfPCell(new Paragraph(NoDecimalMonetaryValue("2"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(laValue);

            // INTEREST RATE / APR
            // Label
            PdfPCell irLabel = new PdfPCell(new Paragraph("Interest Rate / APR:", point10))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(irLabel);
            // Value
            PdfPCell irValue = new PdfPCell(new Paragraph(GetPercentageValue("3") + " / " + GetPercentageValue("799"), point10Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(irValue);

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to cell and remove borders
            PdfPCell cellTable = new PdfPCell(table);

            return cellTable;
        }

        #endregion

        #region Fees, Charges, and Credits Breakdown

        /// <summary>
        /// Creates table to hold the loan charges and borrower
        /// credits breakdown
        /// </summary>
        private void Charges ()
        {
            PdfPTable table = new PdfPTable(3)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] { 48f, 4f, 48f};
            table.SetWidths(columnWidths);            

            table.AddCell(ChargesLeft());

            // Empty cell for formatting
            emptyCell = new PdfPCell(new Paragraph(" "))
            {
                Border = Rectangle.NO_BORDER
            };
            table.AddCell(emptyCell);

            // 
            table.AddCell(ChargesRight());

            table.SpacingAfter = tableSpacingAfter;

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to document
            document.Add(table);  
        } 

        /// <summary>
        /// Creates left column of "Charges" tables
        /// </summary>
        /// <returns></returns>
        private PdfPCell ChargesLeft ()
        {
            PdfPTable table = new PdfPTable(1)
            {
                WidthPercentage = 100
            };
            //float[] columnWidths = new float[] {2f, 1f};
            //table.SetWidths(columnWidths);

            // HEADER
            PdfPCell header = new PdfPCell(new Paragraph("Loan Fees", point10BoldWhite))
            {
                BackgroundColor = headerBackground,
                Border = Rectangle.BOTTOM_BORDER,
                HorizontalAlignment = 1,
                Colspan = 2
            };
            table.AddCell(header);

            table.AddCell(ThirdPartyFees());

            table.AddCell(ThisCompanyFees());

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to cell and remove borders
            PdfPCell cellTable = new PdfPCell(table);

            return cellTable;
        }

        private PdfPCell ThirdPartyFees ()
        {
            PdfPTable table = new PdfPTable(2)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] {2f, 1f};
            table.SetWidths(columnWidths);

            // HEADER - 3rd Party
            PdfPCell header3rd = new PdfPCell(new Paragraph("Paid to 3rd Party", point10Bold))
            {
                Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER,
                HorizontalAlignment = 1,
                Colspan = 2
            };
            table.AddCell(header3rd);

            // All of the "Origination" Fees
            List<KeyValuePair<string, string>> potentialFees = new List<KeyValuePair<string, string>>
            {
	            new KeyValuePair<string, string>("Application Fee","L228"),
	            new KeyValuePair<string, string>("Appraisal Fee","641"),
	            new KeyValuePair<string, string>("Broker Fee","439"),
	            new KeyValuePair<string, string>("City/County Tax Stamps","647"),
	            new KeyValuePair<string, string>("Closing Fee","NEWHUD2.X14"),
	            new KeyValuePair<string, string>("Credit Report","640"),
	            new KeyValuePair<string, string>("Discount Points:","NEWHUD.X1151"),
	            new KeyValuePair<string, string>("Escrow Fee","NEWHUD.X808"),
	            new KeyValuePair<string, string>("Flood Cert","NEWHUD.X400"),
	            new KeyValuePair<string, string>("Lender’s Title Insurance","NEWHUD.X639"),
	            new KeyValuePair<string, string>("Loan Origination Fee: ","454"),
	            new KeyValuePair<string, string>("Origination Credit:","NEWHUD.X1144"),
	            new KeyValuePair<string, string>("Owner’s Title Insurance","NEWHUD.X572"),
	            new KeyValuePair<string, string>("Recording Fees","390"),
	            new KeyValuePair<string, string>("Settlement Fee","NEWHUD2.X11"),
	            new KeyValuePair<string, string>("State Tax Stamps","648"),
	            new KeyValuePair<string, string>("Tax Service Fee","336"),
	            new KeyValuePair<string, string>("Transfer Taxes","NEWHUD.X731"),
	            new KeyValuePair<string, string>("40","41"),
	            new KeyValuePair<string, string>("43","44"),
	            new KeyValuePair<string, string>("348","349"),
	            new KeyValuePair<string, string>("369","370"),
	            new KeyValuePair<string, string>("371","372"),
	            new KeyValuePair<string, string>("373","374"),
	            new KeyValuePair<string, string>("650","644"),
	            new KeyValuePair<string, string>("651","645"),
	            new KeyValuePair<string, string>("931","932"),
	            new KeyValuePair<string, string>("1390","1009"),
	            new KeyValuePair<string, string>("1627","1625"),
	            new KeyValuePair<string, string>("1640","1641"),
	            new KeyValuePair<string, string>("1643","1644"),
	            new KeyValuePair<string, string>("1762","1763"),
	            new KeyValuePair<string, string>("1767","1768"),
	            new KeyValuePair<string, string>("1772","1773"),
	            new KeyValuePair<string, string>("1777","1778"),
	            new KeyValuePair<string, string>("1782","1783"),
	            new KeyValuePair<string, string>("1787","1788"),
	            new KeyValuePair<string, string>("1792","1793"),
	            new KeyValuePair<string, string>("1838","1839"),
	            new KeyValuePair<string, string>("1841","1842"),
	            new KeyValuePair<string, string>("NEWHUD2.X7",  "NEWHUD2.X9"),
	            new KeyValuePair<string, string>("NEWHUD.X1243","NEWHUD.X1245"),
	            new KeyValuePair<string, string>("NEWHUD.X1251","NEWHUD.X1253"),
	            new KeyValuePair<string, string>("NEWHUD.X1259","NEWHUD.X1261"),
	            new KeyValuePair<string, string>("NEWHUD.X126", "NEWHUD.X136"),
	            new KeyValuePair<string, string>("NEWHUD.X1267","NEWHUD.X1269"),
	            new KeyValuePair<string, string>("NEWHUD.X127","NEWHUD.X137"),
	            new KeyValuePair<string, string>("NEWHUD.X1275","NEWHUD.X1277"),
	            new KeyValuePair<string, string>("NEWHUD.X128","NEWHUD.X138"),
	            new KeyValuePair<string, string>("NEWHUD.X1283","NEWHUD.X1285"),
	            new KeyValuePair<string, string>("NEWHUD.X129","NEWHUD.X139"),
	            new KeyValuePair<string, string>("NEWHUD.X1291","NEWHUD.X1293"),
	            new KeyValuePair<string, string>("NEWHUD.X1299","NEWHUD.X1301"),
	            new KeyValuePair<string, string>("NEWHUD.X130","NEWHUD.X140"),
	            new KeyValuePair<string, string>("NEWHUD.X1307","NEWHUD.X1309"),
	            new KeyValuePair<string, string>("NEWHUD.X1315","NEWHUD.X1317"),
	            new KeyValuePair<string, string>("NEWHUD.X1323","NEWHUD.X1325"),
	            new KeyValuePair<string, string>("NEWHUD.x1331","NEWHUD.X1333"),
	            new KeyValuePair<string, string>("NEWHUD.X1339","NEWHUD.X1341"),
	            new KeyValuePair<string, string>("NEWHUD.X1347","NEWHUD.X1349"),
	            new KeyValuePair<string, string>("NEWHUD.X1355","NEWHUD.X1357"),
	            new KeyValuePair<string, string>("NEWHUD.X1363","NEWHUD.X1365"),
	            new KeyValuePair<string, string>("NEWHUD.X1371","NEWHUD.X1373"),
	            new KeyValuePair<string, string>("NEWHUD.X1379","NEWHUD.X1381"),
	            new KeyValuePair<string, string>("NEWHUD.X1387","NEWHUD.X1389"),
	            new KeyValuePair<string, string>("NEWHUD.X1602","NEWHUD.X1604"),
	            new KeyValuePair<string, string>("NEWHUD.X1610","NEWHUD.X1612"),
	            new KeyValuePair<string, string>("NEWHUD.X1618","NEWHUD.X1620"),
	            new KeyValuePair<string, string>("NEWHUD.X1625","NEWHUD.X1627"),
	            new KeyValuePair<string, string>("NEWHUD.X1632","NEWHUD.X1634"),
	            new KeyValuePair<string, string>("NEWHUD.X1640","NEWHUD.X1642"),
	            new KeyValuePair<string, string>("NEWHUD.X1648","NEWHUD.X1650"),
	            new KeyValuePair<string, string>("NEWHUD.X208","NEWHUD.X215"),
	            new KeyValuePair<string, string>("NEWHUD.X209","NEWHUD.X216"),
	            new KeyValuePair<string, string>("NEWHUD.X251","NEWHUD.X254"),
	            new KeyValuePair<string, string>("NEWHUD.X252","NEWHUD.X255"),
	            new KeyValuePair<string, string>("NEWHUD.X253","NEWHUD.X256"),
	            new KeyValuePair<string, string>("NEWHUD.X732","NEWHUD.X733"),
	            new KeyValuePair<string, string>("NEWHUD.X809","NEWHUD.X810"),
	            new KeyValuePair<string, string>("NEWHUD.X811","NEWHUD.X812"),
	            new KeyValuePair<string, string>("NEWHUD.X813","NEWHUD.X814"),
	            new KeyValuePair<string, string>("NEWHUD.X815","NEWHUD.X816"),
	            new KeyValuePair<string, string>("NEWHUD.X951","NEWHUD.X952"),
	            new KeyValuePair<string, string>("NEWHUD.X960","NEWHUD.X961"),
	            new KeyValuePair<string, string>("NEWHUD.X969","NEWHUD.X970"),
	            new KeyValuePair<string, string>("NEWHUD.X978","NEWHUD.X979"),
	            new KeyValuePair<string, string>("NEWHUD.X987","NEWHUD.X988"),
	            new KeyValuePair<string, string>("NEWHUD.X996","NEWHUD.X997"),
            };
            // The actual "Origination" Fees that will appear on the 
            Dictionary<string, string> actualFees = new Dictionary<string, string>();

            // Check potential fees for values
            for (int i = 0; i < potentialFees.Count; i++)
            {
                KeyValuePair<string, string> entry = potentialFees[i];

                // Check for empty string or value of zero
                if (ValueExists(potentialFees[i].Value) == true)
                {
                    // Check for labels not from fields
                    // before adding them to the dictionary.
                    if (i >= 0 && i <= 17)
                    {
                        actualFees.Add(entry.Key, GetMonetaryValue(entry.Value));
                    }
                    else
                        actualFees.Add(GetStringValue(entry.Key), GetMonetaryValue(entry.Value));
                }
            }

            // Alphabetize the fees by placing the keys in a list
            List<string> alphabetizedFees = actualFees.Keys.ToList();
            alphabetizedFees.Sort();

            // Maximum amount of rows for this section.
            int maxRows = 11;

            // Loop through alphabetized fees 
            // and add dictionary values to table
            //foreach (string key in alphabetizedFees)
            for (int i = 0; i < alphabetizedFees.Count; i++)
            {
                // Check if we have space for the fee
                if ((table.Size + 1) >= maxRows)
                {
                    includePage2 = true;

                    // ADDITIONAL CHARGES
                    PdfPCell label = new PdfPCell(new Paragraph("Additional Charges", point10))
                    {
                        Border = Rectangle.LEFT_BORDER,
                        HorizontalAlignment = 0
                    };
                    table.AddCell(label);

                    PdfPCell value = new PdfPCell(new Paragraph("See Page 2", point10Bold))
                    {
                        Border = Rectangle.RIGHT_BORDER,
                        HorizontalAlignment = 2
                    };
                    table.AddCell(value);

                    // Store remaining fees in datatable
                    page2Table = Page2Table();
                    for (int j = i; j < alphabetizedFees.Count; j++)
                    {
                        DataRow row = page2Table.NewRow();

                        row["Label"] = alphabetizedFees[j];
                        row["Value"] = actualFees[alphabetizedFees[j]];

                        page2Table.Rows.Add(row);
                    }
                    break;
                }
                else
                {
                    PdfPCell label = new PdfPCell(new Paragraph(alphabetizedFees[i], point10))
                    {
                        Border = Rectangle.LEFT_BORDER,
                        HorizontalAlignment = 0
                    };
                    table.AddCell(label);

                    PdfPCell value = new PdfPCell(new Paragraph(actualFees[alphabetizedFees[i]], point10Bold))
                    {
                        Border = Rectangle.RIGHT_BORDER,
                        HorizontalAlignment = 2
                    };
                    table.AddCell(value);
                }
            }

            // Count rows and add rows if necessary       
            int usedRows = maxRows - table.Size;
            for (int i = 0; i < usedRows; i++)
            {
                emptyCell = new PdfPCell(new Paragraph(" ", point10))
                {
                    Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER,
                    Colspan = 2
                };
                table.AddCell(emptyCell);
            }

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to cell and remove borders
            PdfPCell cellTable = new PdfPCell(table)
            {
                Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER
            };

            return cellTable;
        }

        private PdfPCell ThisCompanyFees ()
        {
            PdfPTable table = new PdfPTable(2)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] {2f, 1f};
            table.SetWidths(columnWidths);

            // HEADER - On Q
            PdfPCell header3rd = new PdfPCell(new Paragraph("'Company Name' Origination Fees", point10Bold))
            {
                Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER,
                HorizontalAlignment = 1,
                Colspan = 2
            };
            table.AddCell(header3rd);

            // All of the "Origination" Fees
            List<KeyValuePair<string, string>> potentialFees = new List<KeyValuePair<string, string>>
            {
	            new KeyValuePair<string, string>("Processing Fee","1621"),
	            new KeyValuePair<string, string>("Underwriting Fee","367"),
	            new KeyValuePair<string, string>("154","155"),
            };
            // The actual "Origination" Fees that will appear on the 
            Dictionary<string, string> actualFees = new Dictionary<string, string>();

           // Check potential fees for values
            for (int i = 0; i < potentialFees.Count; i++)
            {
                KeyValuePair<string, string> entry = potentialFees[i];

                // Check for empty string or value of zero
                if (ValueExists(potentialFees[i].Value) == true)
                {
                    // Check for labels not from fields
                    // before adding them to the dictionary.
                    if (i >= 0 && i <= 1)
                    {
                        actualFees.Add(entry.Key, GetMonetaryValue(entry.Value));
                    }
                    else
                        actualFees.Add(GetStringValue(entry.Key), GetMonetaryValue(entry.Value));
                }
            }

            // Alphabetize the fees by placing the keys in a list
            List<string> alphabetizedFees = actualFees.Keys.ToList();
            alphabetizedFees.Sort();

            // Loop through alphabetized fees 
            // and add dictionary values to table
            foreach (string key in alphabetizedFees)
            {
                PdfPCell label = new PdfPCell(new Paragraph(key, point10))
                {
                    Border = Rectangle.LEFT_BORDER,
                    HorizontalAlignment = 0
                };
                table.AddCell(label);

                PdfPCell value = new PdfPCell(new Paragraph(actualFees[key], point10Bold))
                {
                    Border = Rectangle.RIGHT_BORDER,
                    HorizontalAlignment = 2
                };
                table.AddCell(value);
            }

            // Count rows and add rows
            // or sum rows if necessary
            int maxRows = 6;
            int usedRows = maxRows - table.Size;
            for (int i = 0; i < usedRows; i++)
            {
                emptyCell = new PdfPCell(new Paragraph(" ", point10))
                {
                    Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER,
                    Colspan = 2
                };
                table.AddCell(emptyCell);
            }

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to cell and remove borders
            PdfPCell cellTable = new PdfPCell(table)
            {
                Border = Rectangle.RIGHT_BORDER | Rectangle.LEFT_BORDER | Rectangle.BOTTOM_BORDER
            };

            return cellTable;
        }

        /// <summary>
        /// Create right column of "Charges" tables
        /// </summary>
        /// <returns></returns>
        private PdfPCell ChargesRight ()
        {
            PdfPTable table = new PdfPTable(2)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] {2f, 1f};
            table.SetWidths(columnWidths);

            // HEADER - Prepaids Section
            PdfPCell headerP = new PdfPCell(new Paragraph("Prepaids", point10BoldWhite))
            {
                BackgroundColor = headerBackground,
                Border = Rectangle.BOTTOM_BORDER,
                HorizontalAlignment = 1,
                Colspan = 2
            };
            table.AddCell(headerP);

            // All of the "Prepaid" Fees
            List<KeyValuePair<string, string>> potentialPrepaids = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("Mortgage Ins. Premium","337"),
                new KeyValuePair<string, string>("Homeowner’s Insurance","642"),
                new KeyValuePair<string, string>("Property Taxes","NEWHUD.X591"),
                new KeyValuePair<string, string>("VA Funding Fee","1050"),
                new KeyValuePair<string, string>("Flood Insurance","643"),
                new KeyValuePair<string, string>("L259","L260"),
                new KeyValuePair<string, string>("1666","1667"),
                new KeyValuePair<string, string>("NEWHUD.X583","NEWHUD.X592"),
                new KeyValuePair<string, string>("NEWHUD.X584","NEWHUD.X592"),
                new KeyValuePair<string, string>("NEWHUD.X1586","NEWHUD.X1588"),
            };

            decimal dailyInt = loan.Fields["332"].ToDecimal();
            // The actual "Prepaids" Fees that will appear on the 
            Dictionary<string, string> prepaidFees = new Dictionary<string, string>
            {
                // This must appear on the form
                { "Daily Int. " + dailyInt.ToString("0.00") + " @ " + GetMonetaryValue("333"), GetMonetaryValue("334") }
            };

            // Check potential fees for values
            for (int i = 0; i < potentialPrepaids.Count; i++)
            {
                KeyValuePair<string, string> entry = potentialPrepaids[i];

                // Check for empty string or value of zero
                if (ValueExists(entry.Value) == true)
                {
                    // Check for labels not from fields
                    // before adding them to the dictionary.
                    if (i >= 0 || i <= 4)
                    {
                        prepaidFees.Add(entry.Key, GetMonetaryValue(entry.Value));
                    }
                    else
                        prepaidFees.Add(GetStringValue(entry.Key), GetMonetaryValue(entry.Value));
                }
            }

            // Alphabetize the fees by placing the keys in a list
            List<string> alphabetizedPrepaids = prepaidFees.Keys.ToList();
            alphabetizedPrepaids.Sort();

            // Loop through alphabetized fees 
            // and add dictionary values to table
            foreach (string key in alphabetizedPrepaids)
            {
                PdfPCell label = new PdfPCell(new Paragraph(key, point10))
                {
                    Border = Rectangle.LEFT_BORDER,
                    HorizontalAlignment = 0
                };
                table.AddCell(label);

                PdfPCell value = new PdfPCell(new Paragraph(prepaidFees[key], point10Bold))
                {
                    Border = Rectangle.RIGHT_BORDER,
                    HorizontalAlignment = 2
                };
                table.AddCell(value);
            }

            // HEADER - Reserves (Escrows)
            PdfPCell headerR = new PdfPCell(new Paragraph("Reserves (Escrows)", point10Bold))
            {
                Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER,
                HorizontalAlignment = 1,
                Colspan = 2
            };
            table.AddCell(headerR);

            // All of the Reserve(Escrows) fees
            List<KeyValuePair<string, string>> potentialReserves = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("Homeowner's Ins "     +       GetStringValue("1387") + " mo. @ " + GetMonetaryValue("230"),   "656"),
                new KeyValuePair<string, string>("Mortgage Ins "        +       GetStringValue("1296") + " mo. @ " + GetMonetaryValue("232"),   "338"),
                new KeyValuePair<string, string>("Property Taxes "      +       GetStringValue("1386") + " mo. @ " + GetMonetaryValue("231"),   "655"),
                new KeyValuePair<string, string>("City Property Tax "   +       GetStringValue("L267") + " mo. @ " + GetMonetaryValue("L268"),  "L269"),
                new KeyValuePair<string, string>("Flood Insurance "     +       GetStringValue("1388") + " mo. @ " + GetMonetaryValue("235"),   "657"),
                new KeyValuePair<string, string>(GetStringValue("1628") + " " + GetStringValue("1629") + " mo. @ " + GetMonetaryValue("1630"),  "1631"),
                new KeyValuePair<string, string>(GetStringValue("660")  + " " + GetStringValue("340")  + " mo. @ " + GetMonetaryValue("253"),   "658"),
                new KeyValuePair<string, string>(GetStringValue("661")  + " " + GetStringValue("341")  + " mo. @ " + GetMonetaryValue("254"),   "659"),
                new KeyValuePair<string, string>("USDA "                +       GetStringValue("NEWHUD.X1706") + " mo. @ " + GetMonetaryValue("NEWHUD.X1707"), "NEWHUD.X1708"),
                new KeyValuePair<string, string>("Aggregate Adjustment", "558")
            };
            // Actaul Reserve(Escrows) fees that will appear on document
            Dictionary<string, string> reserveFees = new Dictionary<string, string>();

            // Check potential fees for values
            for (int i = 0; i < potentialReserves.Count; i++)
            {
                KeyValuePair<string, string> entry = potentialReserves[i];

                // Check for empty string or value of zero
                if (ValueExists(entry.Value) == true)
                {
                    reserveFees.Add(entry.Key, GetMonetaryValue(entry.Value));
                }
            }

            // Alphabetize the fees by placing the keys in a sorted list
            List<string> alphabetizedReserves = reserveFees.Keys.ToList();
            alphabetizedReserves.Sort();

            // Loop through alphabetized fees
            // and add dictionary values to table
            foreach (string key in alphabetizedReserves)
            {
                PdfPCell label = new PdfPCell(new Paragraph(key, point10))
                {
                    Border = Rectangle.LEFT_BORDER,
                    HorizontalAlignment = 0
                };
                table.AddCell(label);

                PdfPCell value = new PdfPCell(new Paragraph(reserveFees[key], point10Bold))
                {
                    Border = Rectangle.RIGHT_BORDER,
                    HorizontalAlignment = 2
                };
                table.AddCell(value);
            }

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to cell and remove borders
            PdfPCell cellTable = new PdfPCell(table);

            return cellTable;
        }

        #endregion

        #region Summary of Estimated Funds Needed to Close

        /// <summary>
        /// Creates Estimated Funds Needed To Close Table
        /// </summary>
        private void EstFundsNeededToClose ()
        {
            PdfPTable table = new PdfPTable(2)
            {
                WidthPercentage = 100
            };

            // HEADER
            PdfPCell header = new PdfPCell(new Paragraph("Estimated Funds Needed to Close", point10BoldWhite))
            {
                BackgroundColor = headerBackground,
                HorizontalAlignment = 1,
                Colspan = 2
            };
            table.AddCell(header);

            table.AddCell(TotalCosts());

            table.AddCell(TotalCredits());

            // CASH TO FROM BORROWER
            // Label
            PdfPCell ctfLabel = new PdfPCell(new Paragraph("Cash [To/From] Borrower (total cost - total credits):", point11Bold))
            {
                Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER,
                HorizontalAlignment = 1,
                Colspan = 2
            };
            table.AddCell(ctfLabel);
            // Value
            PdfPCell ctfValue = new PdfPCell(new Paragraph(GetMonetaryValue("142"), point11))
            {
                Border = Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER | Rectangle.BOTTOM_BORDER,
                HorizontalAlignment = 1,
                Colspan = 2
            };
            table.AddCell(ctfValue);

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to document
            document.Add(table);

            AddFooter(" ");
        }

        /// <summary>
        /// Creates table to hold Total Costs breakdown
        /// housed as cell
        /// </summary>
        /// <returns>Total Costs table</returns>
        private PdfPCell TotalCosts ()
        {
            PdfPTable table = new PdfPTable(2)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] {2f, 1f};
            table.SetWidths(columnWidths);

            // PURCHASE PRICE/PAYOFF
            // Label
            PdfPCell pLabel = new PdfPCell(new Paragraph("Purchase Price/Payoff", point9))
            {
                Border = Rectangle.LEFT_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(pLabel);
            // Value
            PdfPCell pValue = new PdfPCell(new Paragraph(GetMonetaryValue("136"), point9))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(pValue);

            // TOTAL ESTIMATED CLOSING COSTS
            // Label
            PdfPCell ccLabel = new PdfPCell(new Paragraph("Total Estimated Closing Costs", point9))
            {
                Border = Rectangle.LEFT_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(ccLabel);
            // Value
            PdfPCell ccValue = new PdfPCell(new Paragraph(GetMonetaryValue("137"), point9))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(ccValue);

            // TOTAL EST. RESERVES/PREPAID COSTS
            // Label
            PdfPCell rpLabel = new PdfPCell(new Paragraph("Total Est. Reserves/Prepaid Costs", point9))
            {
                Border = Rectangle.LEFT_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(rpLabel);
            // Value
            PdfPCell rpValue = new PdfPCell(new Paragraph(GetMonetaryValue("138"), point9))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(rpValue);

            // DISCOUNT POINTS
            // Label
            PdfPCell dLabel = new PdfPCell(new Paragraph("Discount Points", point9))
            {
                Border = Rectangle.LEFT_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(dLabel);
            // Value
            PdfPCell dValue = new PdfPCell(new Paragraph(GetMonetaryValue("1093"), point9))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(dValue);

            // FHA UFMIP/ VA Funding Fee
            if (ValueExists("1045") == true)
            {
                // Label
                PdfPCell fvLabel = new PdfPCell(new Paragraph("FHA UFMIP/VA Funding Fee", point9))
                {
                    Border = Rectangle.LEFT_BORDER,
                    HorizontalAlignment = 0
                };
                table.AddCell(fvLabel);
                // Value
                PdfPCell fvValue = new PdfPCell(new Paragraph(GetMonetaryValue("969"), point9))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = 2
                };
                table.AddCell(fvValue);
            }

            // Insert empty cell for formatting
            emptyCell = new PdfPCell(new Paragraph(" ", point9))
            {
                Border = Rectangle.RIGHT_BORDER | Rectangle.RIGHT_BORDER,
            };

            // Count rows and add rows
            // or sum rows if necessary
            int maxRows = 10;
            int usedRows = maxRows - table.Size;
            for (int i = 0; i < usedRows; i++)
            {
                emptyCell = new PdfPCell(new Paragraph(" ", point10))
                {
                    Border = Rectangle.LEFT_BORDER,
                    Colspan = 2
                };
                table.AddCell(emptyCell);
            }

            // TOTAL COSTS
            // Label
            PdfPCell tcLabel = new PdfPCell(new Paragraph("Total Costs", point9Bold))
            {
                Border = Rectangle.LEFT_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(tcLabel);
            // Value
            PdfPCell tcValue = new PdfPCell(new Paragraph(GetMonetaryValue("1073"), point9))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(tcValue);

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to cell and remove borders
            PdfPCell cellTable = new PdfPCell(table)
            {
                Border = Rectangle.LEFT_BORDER | Rectangle.BOTTOM_BORDER,
                PaddingRight = estFundPadding
            };

            return cellTable;
        }

        /// <summary>
        /// Creates Total Credits breakdown
        /// housed as cell
        /// </summary>
        /// <returns>Total Credits Table</returns>
        private PdfPCell TotalCredits ()
        {
            PdfPTable table = new PdfPTable(2)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] {2f, 1f};
            table.SetWidths(columnWidths);

            // LOAN AMOUNT
            // Label
            PdfPCell laLabel = new PdfPCell(new Paragraph("Loan Amount", point9))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(laLabel);
            // Value
            PdfPCell laValue = new PdfPCell(new Paragraph(GetMonetaryValue("1109"), point9))
            {
                Border = Rectangle.RIGHT_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(laValue);

            // TOTAL NON-BORROWER PAID CLOSING COSTS
            // Label
            PdfPCell nbcLabel = new PdfPCell(new Paragraph("Total Non-Borrower Paid Closing Costs", point9))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(nbcLabel);
            // Value
            PdfPCell nbcValue = new PdfPCell(new Paragraph(GetMonetaryValue("TNBPCC"), point9))
            {
                Border = Rectangle.RIGHT_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(nbcValue);

            // FHA UFMIP/VA FUNDING FEE FINANCED - Don't Include if value does not exist
            if (ValueExists("1045") == true)
            {
                // Label
                PdfPCell ffLabel = new PdfPCell(new Paragraph("FHA UFMIP/VA Funding Fee Financed", point9))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = 0
                };
                table.AddCell(ffLabel);
                // Value
                PdfPCell ffValue = new PdfPCell(new Paragraph(GetMonetaryValue("1045"), point9))
                {
                    Border = Rectangle.RIGHT_BORDER,
                    HorizontalAlignment = 2
                };
                table.AddCell(ffValue);
            }

            // OTHER - Three fees can exist. Check value of each before adding to table
            List<KeyValuePair<string, string>> potentialOther = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("Other: " + GetStringValue("202"), "141"),
                new KeyValuePair<string, string>("Other: " + GetStringValue("1091"), "1095"),
                new KeyValuePair<string, string>("Other: " + GetStringValue("1106"), "1115")
            };
            Dictionary<string, string> otherFees = new Dictionary<string, string>();

            // Check potential fees for values
            // add them to Dictionary
            for (int i = 0; i < potentialOther.Count; i++)
            {
                KeyValuePair<string, string> entry = potentialOther[i];

                // Check for empty string or value of zero
                if (ValueExists(potentialOther[i].Value) == true)
                {
                    otherFees.Add(entry.Key, GetMonetaryValue(entry.Value));
                }
            }

            // Alphabetize fees
            List<string> abcOther = otherFees.Keys.ToList();
            abcOther.Sort();

            // Loop through alphabetized fees
            // and add them to table
            foreach (string key in abcOther)
            {
                PdfPCell label = new PdfPCell(new Paragraph(key, point9))
                {
                    Border = Rectangle.NO_BORDER,
                    HorizontalAlignment = 0
                };
                table.AddCell(label);

                PdfPCell value = new PdfPCell(new Paragraph(otherFees[key], point9))
                {
                    Border = Rectangle.RIGHT_BORDER,
                    HorizontalAlignment = 2
                };
                table.AddCell(value);
            }

            // FIRST MORTGAGE
            // Label
            PdfPCell fmLabel = new PdfPCell(new Paragraph("First Mortgage", point9))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(fmLabel);
            // Value
            PdfPCell fmValue = new PdfPCell(new Paragraph(GetMonetaryValue("1845"), point9))
            {
                Border = Rectangle.RIGHT_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(fmValue);

            // SUBORDINATE FINANCING/2ND MORTGAGE
            // Label
            PdfPCell sfLabel = new PdfPCell(new Paragraph("Subordinate Financing/2nd Mtg", point9))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(sfLabel);
            // Value
            PdfPCell sfValue = new PdfPCell(new Paragraph(GetMonetaryValue("140"), point9))
            {
                Border = Rectangle.RIGHT_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(sfValue);

            // CLOSING COSTS PAID BY B/L/A/O
            // Label
            PdfPCell ccpbLabel = new PdfPCell(new Paragraph("Closing Costs paid by B/L/A/O", point9))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(ccpbLabel);
            // Value
            PdfPCell ccpbValue = new PdfPCell(new Paragraph(GetMonetaryValue("1852"), point9))
            {
                Border = Rectangle.RIGHT_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(ccpbValue);

            // CLOSING COSTS FROM FIRST LIEN
            // Label
            PdfPCell ccflLabel = new PdfPCell(new Paragraph("Closing Costs from First Lien", point9))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(ccflLabel);
            // Value
            PdfPCell ccflValue = new PdfPCell(new Paragraph(GetMonetaryValue("1851"), point9))
            {
                Border = Rectangle.RIGHT_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(ccflValue);

            // Count rows and add rows
            // or sum rows if necessary
            int maxRows = 10;
            int usedRows = maxRows - table.Size;
            for (int i = 0; i < usedRows; i++)
            {
                emptyCell = new PdfPCell(new Paragraph(" ", point10))
                {
                    Border = Rectangle.RIGHT_BORDER,
                    Colspan = 2
                };
                table.AddCell(emptyCell);
            }

            // TOTAL CREDITS
            // Label
            PdfPCell tcLabel = new PdfPCell(new Paragraph("Total Credits", point9Bold))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = 0
            };
            table.AddCell(tcLabel);
            // Value
            PdfPCell tcValue = new PdfPCell(new Paragraph(GetMonetaryValue("1844"), point9))
            {
                Border = Rectangle.RIGHT_BORDER,
                HorizontalAlignment = 2
            };
            table.AddCell(tcValue);

            // Ensure all rows are drawn
            table.CompleteRow();
            // Add table to cell and remove borders
            PdfPCell cellTable = new PdfPCell(table)
            {
                Border = Rectangle.RIGHT_BORDER | Rectangle.BOTTOM_BORDER,
                PaddingLeft = estFundPadding
            };

            return cellTable;
        }


        #endregion

        /// <summary>
        /// Creates 2nd page which contains all remaining fees not found on page 1
        /// </summary>
        private void AdditionalCharges ()
        {
            document.NewPage();

             PdfPTable headerTable = new PdfPTable(3)
            {
                WidthPercentage = 100
            };
            float[] columnWidths = new float[] { 2f, 3f, 5.5f };
            headerTable.SetWidths(columnWidths);

            // LOGO
            // Get logo from resources and create an iTextSharp Image
            System.Drawing.Bitmap logo = Properties.Resources.Logo;
            Image imageLogo = Image.GetInstance(logo, System.Drawing.Imaging.ImageFormat.Bmp);
            // Create cell and add logo as Element.
            PdfPCell logoCell = new PdfPCell()
            {
                Border = Rectangle.NO_BORDER,
            };
            logoCell.AddElement(imageLogo);
            headerTable.AddCell(logoCell);

            emptyCell = new PdfPCell
            {
                Border = Rectangle.NO_BORDER,
                Colspan = 2
            }; headerTable.AddCell(emptyCell);

            headerTable.CompleteRow();
            headerTable.SpacingAfter = tableSpacingAfter;
            
            document.Add(headerTable);

            // Create table to hold remaining fees
            PdfPTable table = new PdfPTable(2)
            {
                WidthPercentage = 100
            };

            PdfPCell header = new PdfPCell(new Paragraph("3rd Party Loan Fees continued", point10BoldWhite))
            {
                BackgroundColor = headerBackground,
                Colspan = 2
            };
            table.AddCell(header);

            // Add fees to table
            foreach (DataRow row in page2Table.Rows)
            {
                PdfPCell label = new PdfPCell(new Paragraph(row["Label"].ToString(), point10))
                {
                    Border = Rectangle.LEFT_BORDER | Rectangle.BOTTOM_BORDER
                };
                table.AddCell(label);

                PdfPCell value = new PdfPCell(new Paragraph(row["Value"].ToString(), point10))
                {
                    Border = Rectangle.RIGHT_BORDER | Rectangle.BOTTOM_BORDER
                };
                table.AddCell(value);
            }

            table.CompleteRow();
            document.Add(table);

            AddFooter(" ");
        }

        #region Helper Functions

        /// <summary>
        /// Creates a DataTable to hold values that will
        /// appear on Page 2
        /// </summary>
        /// <returns></returns>
        private DataTable Page2Table ()
        {
            DataTable dTable = new DataTable();

            DataColumn colLabel = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Label",
            };
            dTable.Columns.Add(colLabel);

            DataColumn colValue = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Value",
            };
            dTable.Columns.Add(colValue);

            return dTable;
        }

        /// <summary>
        /// Checks Encompass field for value.
        /// Returns true if value exists
        /// </summary>
        /// <param name="fieldID"></param>
        /// <returns></returns>
        private bool ValueExists (String fieldID)
        {
            if (!string.IsNullOrEmpty(loan.Fields[fieldID].ToString()) &&
                float.Parse(loan.Fields[fieldID].ToString()) != 0)
            {
                return true;
            }
            else
                return false;
        }

        /// <summary>
        /// Accepts Encompass Field ID as string, 
        /// gets value, and returns value as string with '$' character
        /// </summary>
        /// <param name="fieldID"></param>
        /// <returns>
        /// Field value with '$' character
        /// or empty string if not value exists
        /// </returns>
        private string GetMonetaryValue (String fieldID)
        {
            if (!string.IsNullOrEmpty(loan.Fields[fieldID].ToString()))
            {
                return "$" + loan.Fields[fieldID].ToString();
            }
            else
                return "    ";
        }

        /// <summary>
        /// Accepts Encompass Field ID as string, 
        /// gets value, and returns value as string with '%' character
        /// </summary>
        /// <param name="fieldID"></param>
        /// <returns>
        /// Field value with '$' character
        /// or empty string if not value exists
        /// </returns>
        private string GetPercentageValue (String fieldID)
        {
            if (!string.IsNullOrEmpty(loan.Fields[fieldID].ToString()))
            {
                return loan.Fields[fieldID].ToString() + "%";
            }
            else
                return "    ";
        }

        /// <summary>
        /// Accepts Encompass Field ID as string, 
        /// gets value, and returns string value
        /// </summary>
        /// <param name="fieldID"></param>
        /// <returns>
        /// Field value with '$' character
        /// or empty string if not value exists
        /// </returns>
        private string GetStringValue (String fieldID)
        {
            if (!string.IsNullOrEmpty(loan.Fields[fieldID].ToString()))
            {
                return loan.Fields[fieldID].ToString();
            }
            else
                return "    ";
        }

        /// <summary>
        /// Rounds the value of an Encompass field to the nearest
        /// whole number
        /// </summary>
        /// <param name="fieldID"></param>
        /// <returns></returns>
        private string NoDecimalMonetaryValue (String fieldID)
        {
            int noDec = (int)Math.Round(loan.Fields[fieldID].ToDecimal());
            string value = "$" + noDec.ToString("N0");
            return value;
        }

        private void AddFooter (string optionalText)
        {
            PdfContentByte cb = writer.DirectContent;
            Rectangle pageSize = document.PageSize;

            // Document Type
            cb.BeginText();
             cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 8.0f);
             cb.SetTextMatrix(pageSize.GetLeft(20), pageSize.GetBottom(15));
             cb.ShowText("Pre-Application Worksheet");
            cb.EndText();

            // Page Number
            string text = "Page " + writer.PageNumber + " " + loan.LoanNumber.ToString();
            cb.BeginText();
            cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, text, pageSize.GetRight(15), pageSize.GetBottom(15), 0);
            cb.EndText();

            // 
            cb.BeginText();
            cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 8.0f);
            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, optionalText,
                pageSize.Width / 2, pageSize.GetBottom(30), 0);
            cb.EndText();
        }
        #endregion

    }
}
