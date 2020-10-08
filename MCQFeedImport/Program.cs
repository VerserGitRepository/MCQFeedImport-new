using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace MCQFeedImport
{
    class Program
    {
        static void Main(string[] args)
        {
            string filepath = System.IO.File.ReadAllText(@"C:\Feed\FileName.txt");
            //Console.WriteLine(filepath);
            if (!string.IsNullOrEmpty(filepath))
            {
                Console.BackgroundColor = ConsoleColor.Gray;
                Console.ResetColor();
                Console.ForegroundColor = ConsoleColor.DarkMagenta;
                Console.WriteLine("------------------Macquarie Feed Import ----------");
                Console.WriteLine();
                Console.ResetColor();
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("####### Please Select Feed Import Options ###########");
                Console.WriteLine();
                Console.ResetColor();
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("[ 1 ]  Import Telstra New Feed");
                Console.WriteLine();
                Console.WriteLine("[ 2 ]  Import Telstra Change Feed.");
                Console.WriteLine();
                Console.WriteLine("[ 3 ]  Import Optus New Feed.");
                Console.WriteLine();
                Console.WriteLine("[ 4 ]  Optus Change Feed");
                Console.WriteLine();
                Console.WriteLine("[ 5 ]   Optus Feed Reprocess");
                Console.WriteLine();
                Console.WriteLine("[ 6 ]   Telstra Feed Reprocess");
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("[ 7 ]   Mail Testing");
                Console.ResetColor();
                Console.WriteLine("#####################################################");
                int option = 0;
                Int32.TryParse(Console.ReadLine(), out option);
                if (option == 1)
                {
                    TelstraNewFeedimport(filepath);
                }
                else if (option == 2)
                {
                    TelstraChangeFeedimport(filepath);
                }
                else if (option == 3)
                {
                    OptusNewFeedimport(filepath);
                }
                else if (option == 4)
                {
                    OptusChangeFeedimport(filepath);
                }
                else if (option == 5)
                {
                    OptusFullFeedReprocess(filepath);
                }
                else if (option == 6)
                {
                    TelstraFullFeedReprocess(filepath);
                }
                else if (option == 7)
                {
                    var responsemodel = new FeedResponseModel {
                        IsFileProcessSuccessful = true,
                        FeedFileName = "Testing feed File name by Basan [Future Feed Processor System Will send Detailed notifation ] ",
                        FeedFileCount = 0
                    };
                    SendMailnotification(responsemodel);
                }
            }
        }

        public static void SendMailnotification(FeedResponseModel responsemodel)
        {
            string templatePath = Path.Combine(@"C:\VerserSourceCodeGitRepo\MCQFeedImport-new\MCQFeedImport\MailTemplate");
            if (true)
            {
                Dictionary<string, string> replacements = new Dictionary<string, string>();
                replacements.Add("FileName", responsemodel.FeedFileName);
                replacements.Add("FeedCount", responsemodel.FeedFileCount.ToString());
                replacements.Add("Date", DateTime.Now.ToString());
                try
                {
                    MailNotificationService.SendMail(String.Format("{0}\\{1}", templatePath, "FeedUpdatenotification.htm"), replacements);
                }
                catch (Exception ex)
                {
                    string LogError = String.Format("Sending email failed {0}", ex.Message);
                }
            }
        }

        public static void TelstraNewFeedimport(string filepath)
        {
            var responsemodel = new FeedResponseModel();
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
                {
                    //create the object for workbook part  
                    WorkbookPart wbPart = doc.WorkbookPart;

                    //statement to get the sheet object  
                    Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);

                    //statement to get the worksheet object by using the sheet id  
                    Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;

                    //Note: worksheet has 8 children and the first child[1] = sheetviewdimension,....child[4]=sheetdata  
                    int wkschildno = 4;

                    //statement to get the sheetdata which contains the rows and cell in table  
                    SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(wkschildno);

                    //List<MCQ_MobileAssetsFeed> assetList = new List<MCQ_MobileAssetsFeed>();
                    int assetrow = 1;
                    for (int row = 1; row < Rows.Count(); row++)
                    {
                        //getting the row as per the specified index of getitem method  
                        Row currentrow = (Row)Rows.ChildElements.GetItem(row);
                        if (!string.IsNullOrEmpty(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart)))
                        {
                            //using (JMSEntities jms = new JMSEntities())
                            using (Blancco BlancoContext = new Blancco())
                            {
                                int assetNo = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart));
                                int serviceNO = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(3), wbPart));
                                string IMEINO = GetCellValue((Cell)currentrow.ChildElements.GetItem(7), wbPart);
                                var assetDB = BlancoContext.MCQFeed
                                    .Where(x => x.AssetNumber == assetNo && x.ServiceNumber == serviceNO && x.IMEI == IMEINO).FirstOrDefault();
                                if (assetDB == null)
                                {
                                    MCQFeed asset = new MCQFeed();
                                    asset.AssetNumber = assetNo;
                                    asset.MCQ_ID = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(1), wbPart));
                                    asset.ServiceNumber = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(3), wbPart));
                                    asset.Swap_Assure_Eligible = GetCellValue((Cell)currentrow.ChildElements.GetItem(4), wbPart);
                                    asset.Make = GetCellValue((Cell)currentrow.ChildElements.GetItem(5), wbPart);
                                    asset.Model = GetCellValue((Cell)currentrow.ChildElements.GetItem(6), wbPart);
                                    asset.IMEI = GetCellValue((Cell)currentrow.ChildElements.GetItem(7), wbPart);
                                    asset.CreatedDate = DateTime.Now;
                                    BlancoContext.MCQFeed.Add(asset);
                                    BlancoContext.SaveChanges();
                                }
                            }
                            Console.WriteLine("No Of Assets processing Count... " + assetrow++);
                        }
                        else
                        {
                            break;
                        }
                    }
                    responsemodel.FeedFileCount = assetrow;
                    responsemodel.FeedFileName = filepath;
                    responsemodel.IsFileProcessSuccessful = true;
                }
            }
            catch (Exception ex)
            {
                responsemodel.IsFileProcessSuccessful = false;
                Console.WriteLine(ex.Message);
            }
            if (responsemodel.IsFileProcessSuccessful)
            {
                SendMailnotification(responsemodel);
            }
        }
        public static void TelstraChangeFeedimport(string filepath)
        {
            var responsemodel = new FeedResponseModel();
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
                {
                    //create the object for workbook part  
                    WorkbookPart wbPart = doc.WorkbookPart;
                    //statement to get the sheet object  
                    Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);
                    //statement to get the worksheet object by using the sheet id  
                    Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;
                    //Note: worksheet has 8 children and the first child[1] = sheetviewdimension,....child[4]=sheetdata  
                    int wkschildno = 4;
                    //statement to get the sheetdata which contains the rows and cell in table  
                    SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(wkschildno);
                    //List<MCQ_MobileAssetsFeed> assetList = new List<MCQ_MobileAssetsFeed>();
                    int assetrow = 1;
                    for (int row = 1; row < Rows.Count(); row++)
                    {
                        //getting the row as per the specified index of getitem method  
                        Row currentrow = (Row)Rows.ChildElements.GetItem(row);
                        if (!string.IsNullOrEmpty(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart)))
                        {
                            //using (JMSEntities jms = new JMSEntities())
                            using (Blancco BlancoContext = new Blancco())
                            {
                                int assetNo = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart));
                                int serviceNO = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(3), wbPart));
                                string IMEINO = GetCellValue((Cell)currentrow.ChildElements.GetItem(7), wbPart);
                                string TelstraMake = GetCellValue((Cell)currentrow.ChildElements.GetItem(5), wbPart);
                                string TelstraModel = GetCellValue((Cell)currentrow.ChildElements.GetItem(6), wbPart);
                                var ChangedServicNumber = BlancoContext.MCQFeed.Where(x => x.AssetNumber == assetNo && x.ServiceNumber != serviceNO
                                ).FirstOrDefault();

                                var ChangedIMEI = BlancoContext.MCQFeed.Where(x => x.AssetNumber == assetNo && x.IMEI != IMEINO).FirstOrDefault();
                                if (ChangedServicNumber != null)
                                {
                                    ChangedServicNumber.ServiceNumber = serviceNO;
                                    ChangedServicNumber.IMEI = IMEINO.ToString();
                                    ChangedServicNumber.Make = TelstraMake;
                                    ChangedServicNumber.Model = TelstraModel;
                                    ChangedServicNumber.UpdatedDate = DateTime.Now;
                                    BlancoContext.SaveChanges();
                                }
                                else
                                {
                                    if (ChangedIMEI != null)
                                    {
                                        ChangedIMEI.ServiceNumber = serviceNO;
                                        ChangedIMEI.IMEI = IMEINO.ToString();
                                        ChangedIMEI.Make = TelstraMake;
                                        ChangedIMEI.Model = TelstraModel;
                                        ChangedIMEI.UpdatedDate = DateTime.Now;
                                        BlancoContext.SaveChanges();
                                    }
                                }
                            }
                            Console.WriteLine("No Of Assets processing Count... " + assetrow++);
                        }
                        else
                        {
                            break;
                        }
                    }
                    responsemodel.FeedFileCount = assetrow;
                    responsemodel.FeedFileName = filepath;
                    responsemodel.IsFileProcessSuccessful = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            if (responsemodel.IsFileProcessSuccessful)
            {
                SendMailnotification(responsemodel);
            }
        }
        public static void OptusNewFeedimport(string filepath)
        {
            var responsemodel = new FeedResponseModel();
            try
            {

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
                {
                    //create the object for workbook part  
                    WorkbookPart wbPart = doc.WorkbookPart;

                    //statement to get the sheet object  
                    Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);

                    //statement to get the worksheet object by using the sheet id  
                    Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;

                    //Note: worksheet has 8 children and the first child[1] = sheetviewdimension,....child[4]=sheetdata  
                    int wkschildno = 4;

                    //statement to get the sheetdata which contains the rows and cell in table  
                    SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(wkschildno);

                    //List<MCQ_MobileAssetsFeed> assetList = new List<MCQ_MobileAssetsFeed>();
                    int assetrow = 1;
                    for (int row = 1; row < Rows.Count(); row++)
                    {
                        //getting the row as per the specified index of getitem method  
                        Row currentrow = (Row)Rows.ChildElements.GetItem(row);
                        if (!string.IsNullOrEmpty(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart)))
                        {
                            //using (JMSEntities jms = new JMSEntities())
                            using (Blancco BlancoContext = new Blancco())
                            {
                                int assetNo = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart));
                                int serviceNO = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(2), wbPart));
                                string IMEINO = GetCellValue((Cell)currentrow.ChildElements.GetItem(5), wbPart);
                                var assetDB = BlancoContext.MCQFeed
                                    .Where(x => x.AssetNumber == assetNo && x.ServiceNumber == serviceNO && x.IMEI == IMEINO).FirstOrDefault();
                                if (assetDB == null)
                                {
                                    MCQFeed asset = new MCQFeed();
                                    asset.AssetNumber = assetNo;
                                    asset.MCQ_ID = 0;
                                    asset.CreatedDate = DateTime.Now;
                                    asset.ServiceNumber = serviceNO;

                                    asset.Make = GetCellValue((Cell)currentrow.ChildElements.GetItem(3), wbPart);
                                    asset.Model = GetCellValue((Cell)currentrow.ChildElements.GetItem(4), wbPart);
                                    asset.IMEI = IMEINO;

                                    BlancoContext.MCQFeed.Add(asset);
                                    BlancoContext.SaveChanges();
                                }
                            }
                            Console.WriteLine("No Of Assets processing Count... " + assetrow++);
                        }
                        else
                        {
                            break;
                        }
                    }
                    responsemodel.FeedFileCount = assetrow;
                    responsemodel.FeedFileName = filepath;
                    responsemodel.IsFileProcessSuccessful = true;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            if (responsemodel.IsFileProcessSuccessful)
            {
                SendMailnotification(responsemodel);
            }
        }
        public static void OptusChangeFeedimport(string filepath)
        {
            var responsemodel = new FeedResponseModel();
            try
            {

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
                {

                    WorkbookPart wbPart = doc.WorkbookPart;
                    Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);
                    Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;
                    int wkschildno = 5;
                    SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(wkschildno);
                    int assetrow = 1;
                    for (int row = 1; row < Rows.Count(); row++)
                    {
                        Row currentrow = (Row)Rows.ChildElements.GetItem(row);
                        if (!string.IsNullOrEmpty(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart)))
                        {

                            using (Blancco BlancoContext = new Blancco())
                            {
                                int assetNo = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart));
                                int serviceNO = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(2), wbPart));
                                string IMEINO = GetCellValue((Cell)currentrow.ChildElements.GetItem(5), wbPart);
                                string OptusMake = GetCellValue((Cell)currentrow.ChildElements.GetItem(3), wbPart);
                                string OptusModel = GetCellValue((Cell)currentrow.ChildElements.GetItem(4), wbPart);

                                var ChangedServicNumber = BlancoContext.MCQFeed
                                    .Where(x => x.AssetNumber == assetNo && x.ServiceNumber != serviceNO).FirstOrDefault();

                                var ChangedIMEI = BlancoContext.MCQFeed.Where(x => x.AssetNumber == assetNo && x.IMEI != IMEINO).FirstOrDefault();
                                if (ChangedServicNumber != null)
                                {
                                    ChangedServicNumber.ServiceNumber = serviceNO;
                                    ChangedServicNumber.IMEI = IMEINO.ToString();
                                    ChangedServicNumber.Make = OptusMake;
                                    ChangedServicNumber.Model = OptusModel;
                                    ChangedServicNumber.UpdatedDate = DateTime.Now;
                                    BlancoContext.SaveChanges();
                                }
                                else
                                {
                                    if (ChangedIMEI != null)
                                    {
                                        ChangedIMEI.ServiceNumber = serviceNO;
                                        ChangedIMEI.IMEI = IMEINO.ToString();
                                        ChangedIMEI.Make = OptusMake;
                                        ChangedIMEI.Model = OptusModel;
                                        ChangedIMEI.UpdatedDate = DateTime.Now;
                                        BlancoContext.SaveChanges();
                                    }
                                }
                            }
                            Console.WriteLine("No Of Assets processing Count... " + assetrow++);
                        }
                        else
                        {
                            break;
                        }
                    }
                    responsemodel.FeedFileCount = assetrow;
                    responsemodel.FeedFileName = filepath;
                    responsemodel.IsFileProcessSuccessful = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            if (responsemodel.IsFileProcessSuccessful)
            {
                SendMailnotification(responsemodel);
            }
        }
        public static void OptusFullFeedReprocess(string filepath)
        {
            var responsemodel = new FeedResponseModel();
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
                {
                    WorkbookPart wbPart = doc.WorkbookPart;
                    Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);
                    Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;
                    int wkschildno = 4;
                    SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(wkschildno);
                    int assetrow = 1;
                    int NewFeedCount = 0;
                    int IMEIFeedUpdatedCount = 0;
                    int ServiceNumberFeedUpdatedCount = 0;
                    for (int row = 1; row < Rows.Count(); row++)
                    {

                        Row currentrow = (Row)Rows.ChildElements.GetItem(row);
                        if (!string.IsNullOrEmpty(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart)))
                        {
                            using (Blancco BlancoContext = new Blancco())
                            {
                                int assetNo = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart));
                                int serviceNO = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(2), wbPart));
                                string IMEINO = GetCellValue((Cell)currentrow.ChildElements.GetItem(5), wbPart);
                                string OptusMake = GetCellValue((Cell)currentrow.ChildElements.GetItem(3), wbPart);
                                string OptusModel = GetCellValue((Cell)currentrow.ChildElements.GetItem(4), wbPart);
                                var ChangedServicNumber = BlancoContext.MCQFeed
                                    .Where(x => x.AssetNumber == assetNo && x.FeedVendor == "Optus" && x.ServiceNumber != serviceNO).FirstOrDefault();

                                var ChangedIMEI = BlancoContext.MCQFeed.Where(x => x.AssetNumber == assetNo && x.FeedVendor == "Optus" && x.IMEI != IMEINO).FirstOrDefault();

                                var _AssetID = BlancoContext.MCQFeed.Where(x => x.AssetNumber == assetNo).FirstOrDefault();

                                if (ChangedServicNumber != null)
                                {
                                    ChangedServicNumber.ServiceNumber = serviceNO;
                                    ChangedServicNumber.IMEI = IMEINO.ToString();
                                    ChangedServicNumber.Make = OptusMake;
                                    ChangedServicNumber.Model = OptusModel;
                                    ChangedServicNumber.UpdatedDate = DateTime.Now;
                                    BlancoContext.SaveChanges();
                                    Console.WriteLine($"Feed Row {assetrow}, Asset ID {ChangedServicNumber.AssetNumber} IMEI Feed Updated..");
                                    ServiceNumberFeedUpdatedCount++;
                                }

                                else if (ChangedIMEI != null)
                                {
                                    ChangedIMEI.ServiceNumber = serviceNO;
                                    ChangedIMEI.IMEI = IMEINO.ToString();
                                    ChangedIMEI.Make = OptusMake;
                                    ChangedIMEI.Model = OptusModel;
                                    ChangedIMEI.UpdatedDate = DateTime.Now;
                                    BlancoContext.SaveChanges();
                                    Console.WriteLine($"Feed Row {assetrow}, Asset ID {ChangedIMEI.AssetNumber} IMEI Feed Updated..");

                                    IMEIFeedUpdatedCount++;
                                }

                                else if (_AssetID == null && ChangedIMEI == null && ChangedServicNumber == null)
                                {
                                    MCQFeed asset = new MCQFeed();
                                    asset.AssetNumber = assetNo;
                                    asset.MCQ_ID = 0;
                                    asset.CreatedDate = DateTime.Now;
                                    asset.ServiceNumber = serviceNO;

                                    asset.Make = GetCellValue((Cell)currentrow.ChildElements.GetItem(3), wbPart);
                                    asset.Model = GetCellValue((Cell)currentrow.ChildElements.GetItem(4), wbPart);
                                    asset.IMEI = IMEINO;
                                    asset.FeedVendor = "Optus";
                                    BlancoContext.MCQFeed.Add(asset);
                                    BlancoContext.SaveChanges();
                                    Console.WriteLine($"Feed Row {assetrow}, Asset ID {_AssetID.AssetNumber} New Feed Creted..");
                                    NewFeedCount++;
                                }
                            }
                            Console.WriteLine("No Of Assets processing Count... " + assetrow++);
                        }
                        else
                        {
                            break;
                        }
                    }
                    string changefeedresp = $"Total Assets Processed {assetrow}, New Feed Created {NewFeedCount}, Service Number Update Count {ServiceNumberFeedUpdatedCount} and IMEI Update Count {IMEIFeedUpdatedCount}";

                    responsemodel.FeedFileCount = assetrow;
                    responsemodel.FeedFileName = filepath + changefeedresp;
                    responsemodel.IsFileProcessSuccessful = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            if (responsemodel.IsFileProcessSuccessful)
            {
                SendMailnotification(responsemodel);
            }
        }
        public static void TelstraFullFeedReprocess(string filepath)
        {
            var responsemodel = new FeedResponseModel();
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
                {
                    WorkbookPart wbPart = doc.WorkbookPart;
                    Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);
                    Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;
                    int wkschildno = 4;
                    SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(wkschildno);
                    int assetrow = 1;
                    int NewFeedCount = 0;
                    int IMEIFeedUpdatedCount = 0;
                    int ServiceNumberFeedUpdatedCount = 0;
                    for (int row = 1; row < Rows.Count(); row++)
                    {

                        Row currentrow = (Row)Rows.ChildElements.GetItem(row);
                        if (!string.IsNullOrEmpty(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart)))
                        {
                            using (Blancco BlancoContext = new Blancco())
                            {
                                int assetNo = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(0), wbPart));
                                int serviceNO = Convert.ToInt32(GetCellValue((Cell)currentrow.ChildElements.GetItem(3), wbPart));
                                string IMEINO = GetCellValue((Cell)currentrow.ChildElements.GetItem(7), wbPart);
                                string TelstraMake = GetCellValue((Cell)currentrow.ChildElements.GetItem(5), wbPart);
                                string TelstraModel = GetCellValue((Cell)currentrow.ChildElements.GetItem(6), wbPart);


                                var ChangedServicNumber = BlancoContext.MCQFeed
                                    .Where(x => x.AssetNumber == assetNo && x.FeedVendor == "Telstra" && x.ServiceNumber != serviceNO).FirstOrDefault();

                                var ChangedIMEI = BlancoContext.MCQFeed.Where(x => x.AssetNumber == assetNo && x.FeedVendor == "Telstra" && x.IMEI != IMEINO).FirstOrDefault();
                                var _AssetID = BlancoContext.MCQFeed.Where(x => x.AssetNumber == assetNo).FirstOrDefault();
                                if (ChangedServicNumber != null)
                                {
                                    ChangedServicNumber.ServiceNumber = serviceNO;
                                    ChangedServicNumber.IMEI = IMEINO.ToString();
                                    ChangedServicNumber.Make = TelstraMake;
                                    ChangedServicNumber.Model = TelstraModel;
                                    ChangedServicNumber.UpdatedDate = DateTime.Now;
                                    BlancoContext.SaveChanges();
                                    Console.WriteLine($"Feed Row {assetrow}, Asset ID {ChangedServicNumber.AssetNumber} IMEI Feed Updated..");
                                    ServiceNumberFeedUpdatedCount++;
                                }

                                else if (ChangedIMEI != null)
                                {
                                    ChangedIMEI.ServiceNumber = serviceNO;
                                    ChangedIMEI.IMEI = IMEINO.ToString();
                                    ChangedIMEI.Make = TelstraMake;
                                    ChangedIMEI.Model = TelstraModel;
                                    ChangedIMEI.UpdatedDate = DateTime.Now;
                                    BlancoContext.SaveChanges();
                                    Console.WriteLine($"Feed Row {assetrow}, Asset ID {ChangedIMEI.AssetNumber} IMEI Feed Updated..");

                                    IMEIFeedUpdatedCount++;
                                }

                                else if (_AssetID == null && ChangedIMEI == null && ChangedServicNumber == null)
                                {
                                    MCQFeed asset = new MCQFeed();
                                    asset.AssetNumber = assetNo;
                                    //asset.MCQ_ID = 0;
                                    asset.CreatedDate = DateTime.Now;
                                    asset.ServiceNumber = serviceNO;
                                    asset.Make = TelstraMake; //GetCellValue((Cell)currentrow.ChildElements.GetItem(3), wbPart);
                                    asset.Model = TelstraModel;// GetCellValue((Cell)currentrow.ChildElements.GetItem(4), wbPart);
                                    asset.IMEI = IMEINO;
                                    asset.FeedVendor = "Telstra";
                                    BlancoContext.MCQFeed.Add(asset);
                                    BlancoContext.SaveChanges();
                                    Console.WriteLine($"Feed Row {assetrow}, Asset ID {_AssetID.AssetNumber} New Feed Creted..");
                                    NewFeedCount++;
                                }
                            }
                            Console.WriteLine("No Of Assets processing Count... " + assetrow++);
                        }
                        else
                        {
                            break;
                        }
                    }
                    string changefeedresp = $"Total Assets Processed {assetrow}, New Feed Created {NewFeedCount}, Service Number Update Count {ServiceNumberFeedUpdatedCount} and IMEI Update Count {IMEIFeedUpdatedCount}";
                    responsemodel.FeedFileCount = assetrow;
                    responsemodel.FeedFileName = filepath + changefeedresp;
                    responsemodel.IsFileProcessSuccessful = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            if (responsemodel.IsFileProcessSuccessful)
            {
                SendMailnotification(responsemodel);
            }
        }
        private static string GetCellValue(Cell theCell, WorkbookPart wbPart)
        {
            string value = "";
            if (theCell != null)
            {
                value = theCell.InnerText;

                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable =
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (stringTable != null)
                            {
                                value =
                                    stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }
            return value;
        }

    }
}

