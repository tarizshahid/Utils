using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Globalization;
using OpenQA.Selenium.Support.Extensions;
using System.Text.RegularExpressions;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager;
using System.IO;
using OpenQA.Selenium.DevTools;
using Serilog.Events;
using Log = Serilog.Log;
using Serilog;
using NReco.PdfGenerator;
using OpenQA.Selenium.Support.UI;
using PBS.Entities;
using PBS.DAL;
using RPASuiteDataService.DBEntities;
using System.Net;
using WebDriverManager.Helpers;

namespace PBS
{
    class BookingSheet
    {
        public static List<int> SummaryID = new List<int>();
        public static int gLogId = 0;
        public static DateTime StartTime = DateTime.Now;
        static string URL = @"https://flddacapp.eclinicalweb.com/mobiledoc/jsp/webemr/index.jsp#/mobiledoc/jsp/webemr/jellybean/officevisit/officeVisits.jsp";

        static async Task Main(string[] args)
        {
            try
            {
                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Debug()
                    .WriteTo.File("BookingSheetLog.txt")
                    .WriteTo.Console(restrictedToMinimumLevel: LogEventLevel.Information)
                    .CreateLogger();

                bool isUrlWorking = await CheckUrl();
                if (isUrlWorking == false)
                {
                    Log.Information("URL not Working.");
                    Log.Information("Exiting Application.");
                    PBSDataService dsObj = new();
                    var emailmessage = $"The following URL {URL} is not working for DDA Booking Sheet.\n";
                    var emailsubject = "URL not working : Booking Sheet DDA";
                    dsObj.EmailNotification(emailmessage, emailsubject);
                    gLogId = PBSDataService.SaveGeneralLog(new GeneralLog
                    {
                        DataListMappingID = 110,
                        StartTime = StartTime,
                        Status = "URL not working",
                        EndTime = null,
                    });
                    var ip = Dns.GetHostEntry(Dns.GetHostName())
                        .AddressList.First(address => address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                        .ToString();
                    PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                    {
                        GeneralLogID = gLogId,
                        EventTime = DateTime.Now,
                        ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                        CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                        SubCategoryName = "Booking Sheet",
                        EventDescription = $"URL not Working : {URL} on IP Address : {ip} \t Machine Name : {Environment.MachineName} \t Request Time : {DateTime.Now}",
                    });

                    PBSDataService.SaveGeneralLog(new GeneralLog
                    {
                        ID = gLogId,
                        DataListMappingID = 110,
                        StartTime = StartTime,
                        Status = "URL not working",
                        EndTime = DateTime.Now,
                    });
                    Environment.Exit(0);
                }

                //checks PBS Process request dates (if 0 then continue)
                var CheckDatesList = PBSDataService.getPBSProcRequest(-1);
                var Date20days = DateTime.Now.AddDays(20);
                bool exists = CheckDatesList.Any(item => item.ProcessDate.HasValue && item.ProcessDate.Value.Date == Date20days.Date);
                
                //Add-ons implmentation
                var AddOnDate = DateTime.Now.Date.AddDays(2);
                PBSProcessRequests? DateToUpdate = CheckDatesList.FirstOrDefault(checkDate => checkDate.ProcessDate == AddOnDate);
                if (DateToUpdate.isProcessed == 1 && DateToUpdate.UpdatedDate < DateTime.Now.Date.AddDays(-5))
                {
                    DateToUpdate.isProcessed = 0;
                    PBSDataService.savePBSProcRequest(DateToUpdate);
                    Log.Information($"Add-ons Date : {DateToUpdate.ProcessDate} Added.");
                }

                if (!exists)
                {
                    //20 Days ahead addition
                    PBSProcessRequests procreq = new PBSProcessRequests()
                    {
                        PracticeID = 6,
                        ProcessDate = Date20days,
                        isProcessed = 0,
                        ID = 0
                    };
                    PBSDataService.savePBSProcRequest(procreq);
                }

                var toProcess = PBSDataService.getPBSProcRequest(0);
                if (toProcess.Count > 0)
                {
                    Log.Information($"{toProcess.Count} Date(s) avaialable to process");
                    foreach (var item in toProcess)
                    {
                        string? date = item.ProcessDate?.ToString("MM/dd/yyyy");
                        Log.Information($"Application Started for {item.ProcessDate?.ToString("MM/dd/yyyy")}");
                        if (Status(date) == 1)
                        {
                            PBSDataService.FillPDF(date);
                            var excelpath = PBSDataService.ExcelSheetCreationPlusSending(date); //MM-dd-yyyy
                            if (excelpath != "ZeroPatients")
                            {
                                PBSDataService.SendFiles($"Here's the booking sheet Excel & PDF(s) ZIP File for {date}", $"Patient Booking Sheet {date}", excelpath, date);
                                item.isProcessed += 1;
                                PBSDataService.savePBSProcRequest(item);
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = $"isProcessed updated for {date}",
                                });
                                PBSDataService.SaveGeneralLog(new GeneralLog
                                {
                                    ID = gLogId,
                                    DataListMappingID = 110,
                                    StartTime = StartTime,
                                    Status = "Completed",
                                    EndTime = DateTime.Now
                                });
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = "Application Ended Sucessfully",
                                });
                                Log.Information($"isProcessed updated for {date}");
                                Environment.Exit(0);
                            }
                            else
                            {
                                item.isProcessed += 1;
                                PBSDataService.savePBSProcRequest(item);
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = $"isProcessed updated for {date}",
                                });
                                PBSDataService.SaveGeneralLog(new GeneralLog
                                {
                                    ID = gLogId,
                                    DataListMappingID = 110,
                                    StartTime = StartTime,
                                    Status = "Completed",
                                    EndTime = DateTime.Now
                                });
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "PMSCode : DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = "Application Ended Sucessfully",
                                });
                                Log.Information($"isProcessed updated for {date}");
                                Environment.Exit(0);
                            }
                        }
                        if (Status(date) == 0)
                        {
                            Status(date);
                        }
                    }
                    Log.Information("Application Ended Sucessfully");
                    Environment.Exit(0);
                }
                else
                {
                    Log.Information("No Date(s) available to process.");
                    Environment.Exit(0);
                }
            }
            catch (Exception ex)
            {
                PBSDataService dsObj = new();
                var emailmessage = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}";
                var emailsubject = "Exception Occured : Booking Sheet DDA";
                dsObj.EmailNotification(emailmessage, emailsubject);
            }
        }

        //main function to start processing
        public static int Status(string date)
        {
            gLogId = PBSDataService.SaveGeneralLog(new GeneralLog
            {
                DataListMappingID = 110,
                StartTime = StartTime,
                Status = "Started",
                EndTime = null
            });

            int Success = 0;
            var Currentdate = DateTime.Now.ToString("MM/dd/yyyy");
            CultureInfo provider = CultureInfo.InvariantCulture;

            //check date (should not be less than today)
            if (DateTime.Compare(Convert.ToDateTime(date), Convert.ToDateTime(Currentdate)) < 0)
            {
                Log.Information("Date can not be less than current date\nApplication Ended");
                Environment.Exit(0); //Modification required
            }
            PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
            {
                GeneralLogID = gLogId,
                EventTime = DateTime.Now,
                ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                SubCategoryName = "Booking Sheet",
                EventDescription = $"Booking Sheet process started for {date}",
            });
            PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
            {
                GeneralLogID = gLogId,
                EventTime = DateTime.Now,
                ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                SubCategoryName = "Booking Sheet",
                EventDescription = $"Navigating to URL : {URL}",
            });

            var chromeOptions = new ChromeOptions();
            chromeOptions.AddUserProfilePreference("download.prompt_for_download", false);
            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
            chromeOptions.AddExcludedArgument("disable-popup-blocking");

            ///set browser headleass
            //chromeOptions.AddArguments("--headless");
            chromeOptions.AddArguments("--disable-gpu");
            chromeOptions.AddArguments("--window-size=1920,1080");
            chromeOptions.AddArguments("--allow-insecure-localhost");

            var chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;    // hide the console.
            Log.Information("Initializing WebDriver");
            new DriverManager().SetUpDriver(new ChromeConfig(), VersionResolveStrategy.MatchingBrowser); //install matching version webdriver


            var timespan = TimeSpan.FromMinutes(5);
            IWebDriver driver = new ChromeDriver(chromeDriverService, chromeOptions, timespan);
            Log.Information("WebDriver Initalized");
            PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
            {
                GeneralLogID = gLogId,
                EventTime = DateTime.Now,
                ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                SubCategoryName = "Booking Sheet",
                EventDescription = $"WebDriver Initalized",
            });
            int count = 0;

            try
            {
                Log.Information("Login to Website started");
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
                driver.Navigate().GoToUrl(URL);
                Thread.Sleep(1000);
                driver.Manage().Window.Maximize();


                //Get username/password from appsettings for eCW
                var credentials = PBSDataService.GetCredentials();
                //Login Page
                Log.Information("Entering Username/Password");
                driver.FindElement(By.XPath("//*[@id='doctorID']")).SendKeys(credentials.username);
                driver.FindElement(By.XPath("//*[@id='doctorID']")).SendKeys(Keys.Enter);
                Thread.Sleep(5000);
                driver.FindElement(By.XPath("//*[@id='passwordField']")).SendKeys(credentials.password);
                driver.FindElement(By.XPath("//*[@id='Login']")).Click();

                //redirecting to Office visits
                driver.Navigate().GoToUrl(URL);
                Thread.Sleep(5000);

                //facility pop closure
                try
                {
                    driver.FindElement(By.XPath("//*[@id='closeID']")).Click();
                }
                catch (Exception ex)
                {
                    Log.Information("Close Popup not found.");
                    goto cont;
                }

            cont:
                Log.Information("Logged in sucessfully");
                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                {
                    GeneralLogID = gLogId,
                    EventTime = DateTime.Now,
                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                    SubCategoryName = "Booking Sheet",
                    EventDescription = $"Logged in Successfully",
                });
                //Date
                Thread.Sleep(2000);
                //Date read only attribute removal
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("document.getElementById('officeVisitsIpt3').removeAttribute('readonly', 0);");
                Thread.Sleep(3000);
                driver.FindElement(By.XPath("//*[@id='officeVisitsIpt3']")).Clear(); //Date set
                Thread.Sleep(3000);
                Log.Information("Setting Date : " + date);
                driver.FindElement(By.XPath("//*[@id='officeVisitsIpt3']")).SendKeys(date);
                Thread.Sleep(3000);
                driver.FindElement(By.XPath("//*[@id='officeVisitsIpt3']")).SendKeys(Keys.Enter);
                Thread.Sleep(3000);

                var procedureDate = date;
                var ProviderName = "";
                var FacilityName = "";

                //Main Loop ,starts with Provider 
                for (int k = 0; k <= 1; k++)
                {
                    Log.Information("Iterating through Provider");
                    js.ExecuteScript("window.onbeforeunload = function() {};");
                    Thread.Sleep(8000);
                    driver.FindElement(By.XPath("/html/body/div[3]/div[4]/section/div/div[2]/section/div[1]/div[2]/div/div[1]/div/form/table/tbody/tr/td[1]/div[1]/div[1]/div/div/div/div/form/div/div/div/div/button")).Click(); //lookup button provider
                    Thread.Sleep(3000);
                    ProviderName = driver.FindElement(By.XPath($"//*[@id='provider-lookupLink1ngR{k}']")).Text.ToString();
                    driver.FindElement(By.XPath($"//*[@id='provider-lookupLink1ngR{k}']")).Click(); //provider1 : Ahmed,Saeed
                    Thread.Sleep(3000);

                    var datecompare = driver.FindElement(By.XPath("//*[@id='officeVisitsIpt3']")).GetAttribute("value");
                    if (date != datecompare)
                    {
                        Log.Information("Date Input Error\tRestarting Application");
                        driver.Quit();
                        Console.Clear();

                        //Recursive call when date input error occurs
                        Status(date);
                    }

                    //Facilites List
                    List<string> Facilities = new List<string>
                    {
                        "//*[@id='facility-lookupLink1ngR3']",
                        "//a[contains(text(), 'Ambulatory Surgery Center')]",
                        "//a[contains(text(), 'TMIS')]"
                    };

                    for (int il = 0; il <= 2; il++)
                    {
                        Log.Information("Iterating through Facilities");
                        Thread.Sleep(2000);
                        driver.FindElement(By.XPath("/html/body/div[3]/div[4]/section/div/div[2]/section/div[1]/div[2]/div/div[1]/div/form/table/tbody/tr/td[1]/div[2]/div[2]/div/div/div/form/div/div/div/div/button")).Click(); //lookup button faciltiy
                        Thread.Sleep(3000);
                        FacilityName = driver.FindElement(By.XPath($"{Facilities[il]}")).Text.ToString();
                        driver.FindElement(By.XPath($"{Facilities[il]}")).Click();
                        Thread.Sleep(1000);
                        driver.FindElement(By.XPath("//*[@id='officeVisitsBtn16']")).Click(); //lookup
                        Thread.Sleep(3000);

                        //Check patient count , bottom right of screen
                        var totalCount = Convert.ToInt32(driver.FindElement(By.XPath("//*[@id='open']/div[5]/div[2]/div-pagination-control/label/span[2]")).Text);
                        if (totalCount < 1)
                        {
                            PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                            {
                                GeneralLogID = gLogId,
                                EventTime = DateTime.Now,
                                ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                SubCategoryName = "Booking Sheet",
                                EventDescription = $"No Office Visits found against {date}\t|\tProvider : {ProviderName}\t|\tFacility : {FacilityName}",
                            });
                            Log.Information($"No Office Visits found against {date}\t|\tProvider : {ProviderName}\t|\tFacility : {FacilityName}");

                            var checkNullPatinDB = PBSDataService.getBookingSheetSummary(procedureDate, ProviderName, FacilityName);
                            if (checkNullPatinDB.Count == 0)
                            {
                                PBSDataService.saveBookingSheetSummary(procedureDate, ProviderName, FacilityName, 0, 0, 0);
                                continue;
                            }
                            else
                            {
                                continue;
                            }
                        }

                        PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                        {
                            GeneralLogID = gLogId,
                            EventTime = DateTime.Now,
                            ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                            CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                            SubCategoryName = "Booking Sheet",
                            EventDescription = $"{totalCount} Office Visits found against {date}\t|\tProvider : {ProviderName}\t|\tFacility : {FacilityName}",
                        });
                        Log.Information($"{totalCount} Office Visits found against {date}\t|\tProvider : {ProviderName}\t|\tFacility : {FacilityName}");

                        var visitType = "";
                        List<string> patientsXpaths = new List<string>();
                        List<string> ApptTime = new List<string>();
                        List<string> PatientReason = new List<string>();
                        var CheckPatients = new PatientInfo();
                        int AlreadyinDatabase = 0;
                        int totalProcedurePatients = 0;
                        Log.Information($"Checking already processed patients.");
                        PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                        {
                            GeneralLogID = gLogId,
                            EventTime = DateTime.Now,
                            ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                            CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                            SubCategoryName = "Booking Sheet",
                            EventDescription = $"Checking already processed patients.",
                        });
                        for (int PatChk = 1; PatChk <= totalCount; PatChk++)
                        {
                            visitType = driver.FindElement(By.XPath($"//*[@id='this-table']/tbody/tr[{PatChk}]/td[3]")).Text.ToString(); // visit type should be procedure
                            if (visitType.Trim() == "Procedure" || visitType.Trim() == "AHWC")
                            {
                                totalProcedurePatients++;
                                Thread.Sleep(100);
                                var patientName = driver.FindElement(By.XPath($"//*[@id='this-table']/tbody/tr[{PatChk}]/td[5]")).Text.Split(',');
                                Thread.Sleep(100);
                                CheckPatients.PatientLastName = patientName[0].Trim();
                                try
                                {
                                    if (patientName[1].Trim().Split(' ')[1].Length == 1)
                                    {
                                        CheckPatients.PatientFirstName = patientName[1].Trim().Split(' ')[0];
                                    }
                                    if (patientName[1].Trim().Split(' ')[1].Length > 1)
                                    {
                                        CheckPatients.PatientFirstName = patientName[1].Trim();
                                    }
                                }
                                catch (Exception)
                                {
                                    CheckPatients.PatientFirstName = patientName[1].Trim();
                                    goto nextstatement;
                                }
                            nextstatement:
                                CheckPatients.ProcedureDate = date;
                                Thread.Sleep(100);
                                CheckPatients.Reason = driver.FindElement(By.XPath($"//*[@id='this-table']/tbody/tr[{PatChk}]/td[7]")).Text.ToString();
                                Thread.Sleep(100);
                                CheckPatients.ProcedureTime = driver.FindElement(By.XPath($"//*[@id='this-table']/tbody/tr[{PatChk}]/td[4]")).Text.ToString();
                                Thread.Sleep(100);
                                if (PBSDataService.CheckPatientinDB(CheckPatients.ProcedureDate, FacilityName, CheckPatients.ProcedureTime, CheckPatients.Reason, CheckPatients.PatientLastName, CheckPatients.PatientFirstName, ProviderName) == false)
                                {
                                    AlreadyinDatabase++;
                                    patientsXpaths.Add($"//*[@id='this-table']/tbody/tr[{PatChk}]/td[5]");
                                    ApptTime.Add($"//*[@id='this-table']/tbody/tr[{PatChk}]/td[4]");
                                    PatientReason.Add(driver.FindElement(By.XPath($"//*[@id='this-table']/tbody/tr[{PatChk}]/td[7]")).Text);
                                }
                                Thread.Sleep(100);
                            }
                        }
                        PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                        {
                            GeneralLogID = gLogId,
                            EventTime = DateTime.Now,
                            ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                            CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                            SubCategoryName = "Booking Sheet",
                            EventDescription = $"Total Patients for Procedure : " + totalProcedurePatients,
                        });
                        Log.Information($"Total Patients for Procedure : " + totalProcedurePatients);
                        PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                        {
                            GeneralLogID = gLogId,
                            EventTime = DateTime.Now,
                            ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                            CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                            SubCategoryName = "Booking Sheet",
                            EventDescription = $"{totalProcedurePatients - AlreadyinDatabase} patients already in Database.",
                        });
                        Log.Information($"{totalProcedurePatients - AlreadyinDatabase} patients already in Database.");
                        Log.Information($"Checking Addon patients.");

                        var Summaryobj = PBSDataService.getBookingSheetSummary(procedureDate, ProviderName, FacilityName);
                        if (Summaryobj.FirstOrDefault()?.ID != 0 && (totalProcedurePatients - Summaryobj.FirstOrDefault()?.TotalPatients > 0))
                        {
                            SummaryID.Add(PBSDataService.saveBookingSheetSummary(procedureDate, ProviderName, FacilityName, Summaryobj.FirstOrDefault().isEmailSent, totalProcedurePatients, (totalProcedurePatients - Summaryobj.FirstOrDefault().TotalPatients), Summaryobj.FirstOrDefault().ID));
                            Log.Information($"{(totalProcedurePatients - Summaryobj.FirstOrDefault()?.TotalPatients)} Addon Patients found.");
                            PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                            {
                                GeneralLogID = gLogId,
                                EventTime = DateTime.Now,
                                ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                SubCategoryName = "Booking Sheet",
                                EventDescription = $"{(totalProcedurePatients - Summaryobj.FirstOrDefault()?.TotalPatients)} Addon Patients found.",
                            });
                        }
                        else if (Summaryobj?.FirstOrDefault()?.TotalPatients != (totalProcedurePatients - AlreadyinDatabase) && Summaryobj.Count != 0)
                        {
                            Log.Information("Zero Addon Patients found");
                            SummaryID.Add(PBSDataService.saveBookingSheetSummary(procedureDate, ProviderName, FacilityName, Summaryobj.FirstOrDefault().isEmailSent, totalProcedurePatients, (totalProcedurePatients - Summaryobj.FirstOrDefault().TotalPatients), Summaryobj.FirstOrDefault().ID));
                        }
                        else if (Summaryobj?.FirstOrDefault()?.TotalPatients == (totalProcedurePatients - AlreadyinDatabase) && Summaryobj?.FirstOrDefault()?.isEmailSent > 0)
                        {
                            Log.Information("Zero Addon Patients found");
                            SummaryID.Add(PBSDataService.saveBookingSheetSummary(procedureDate, ProviderName, FacilityName, Summaryobj.FirstOrDefault().isEmailSent, totalProcedurePatients, 0, Summaryobj.FirstOrDefault().ID));
                        }
                        else
                        {
                            Log.Information("0 Addon Patients found");
                            SummaryID.Add(PBSDataService.saveBookingSheetSummary(procedureDate, ProviderName, FacilityName, 0, totalProcedurePatients, 0));
                        }

                        if (patientsXpaths.Count > 0)
                        {
                            Log.Information($"Retrieving remainig {patientsXpaths.Count} patients");
                        }
                        for (int it = 0; it < patientsXpaths.Count; it++)
                        {
                        oneprogressnote:
                            var procedureTime = driver.FindElement(By.XPath(ApptTime[it])).Text.ToString();
                            var Reason = PatientReason[it];
                            Log.Information($"Started processing patient : {driver.FindElement(By.XPath(patientsXpaths[it])).Text} for {Reason} | {it + 1}/{patientsXpaths.Count}");
                            Thread.Sleep(1000);
                            PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                            {
                                GeneralLogID = gLogId,
                                EventTime = DateTime.Now,
                                ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                SubCategoryName = "Booking Sheet",
                                EventDescription = $"Started processing patient : {driver.FindElement(By.XPath(patientsXpaths[it])).Text} for {Reason} | {it + 1}/{patientsXpaths.Count}",
                            });
                            try
                            {
                                driver.FindElement(By.XPath("//*[contains(text(), 'Total Counts')]")).Click();
                            }
                            catch (Exception ex)
                            {
                                Log.Information($"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                                goto ctd3;
                            }
                        ctd3:

                            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", driver.FindElement(By.XPath(patientsXpaths[it])));
                            driver.FindElement(By.XPath(patientsXpaths[it])).Click();
                            Thread.Sleep(10000);
                            driver.Navigate().Refresh();
                            try
                            {
                                driver.SwitchTo().Alert().Accept();
                            }
                            catch (Exception ex)
                            {
                                Log.Information($"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                                goto ctd2;
                            }
                        ctd2:
                            Thread.Sleep(8000);
                            driver.FindElement(By.XPath("//*[@id='encDropDown']/div[1]")).Click();
                            List<string> ProgressNotesDates = new List<string>();

                            try
                            {
                                for (int i = 1; i < 50; i++)
                                {
                                    Thread.Sleep(500);
                                    var Visitlock = Convert.ToBoolean(driver.ExecuteJavaScript<object>($"return document.querySelector('#encDropDownList > li:nth-child({i}) > i').classList.contains('icon-locked-physical-visit-progressnote')"));
                                    Thread.Sleep(500);
                                    var Visit = Convert.ToBoolean(driver.ExecuteJavaScript<object>($"return document.querySelector('#encDropDownList > li:nth-child({i}) > i').classList.contains('icon-physical-visit-progressnote')"));
                                    Thread.Sleep(500);

                                    if (Visitlock == true)
                                    {
                                        var extractDate = Convert.ToString(driver.ExecuteJavaScript<object>($"return document.querySelector('#encDropDownList > li:nth-child({i})').textContent"));
                                        Thread.Sleep(1000);
                                        var VisitDate = extractDate?.Split(' ');
                                        Thread.Sleep(1000);
                                        ProgressNotesDates.Add(VisitDate[0]);
                                    }
                                    else if (Visit == true)
                                    {
                                        var extractDate = Convert.ToString(driver.ExecuteJavaScript<object>($"return document.querySelector('#encDropDownList > li:nth-child({i})').textContent"));
                                        Thread.Sleep(1000);
                                        var VisitDate = extractDate?.Split(' ');
                                        Thread.Sleep(1000);
                                        ProgressNotesDates.Add(VisitDate[0]);
                                    }
                                    else
                                    {
                                        ProgressNotesDates.Add("Telephone/Web");
                                    }
                                }
                            }
                            catch (JavaScriptException)
                            {
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = $"JavaScript Exception occured, Visit is either Web or Telephone etc",
                                });
                                Log.Information($"JavaScript Exception occured, Visit is either Web or Telephone etc");
                                goto endloop;
                            }
                            catch (IndexOutOfRangeException ex)
                            {
                                Log.Information($"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                                goto endloop;
                            }
                        endloop:
                            int j = 0;
                            int progressNoteIteratorValue = 0;
                            while (j < 50)
                            {
                                if (ProgressNotesDates[j] != "Telephone/Web" && DateTime.Now.Date < Convert.ToDateTime(ProgressNotesDates[j]) && ProgressNotesDates.Count == 1 || (ProgressNotesDates.Count == 1 && ProgressNotesDates[j] == "Telephone/Web"))
                                {
                                    driver.Navigate().GoToUrl(URL);
                                    Thread.Sleep(3000);
                                    PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                    {
                                        GeneralLogID = gLogId,
                                        EventTime = DateTime.Now,
                                        ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                        CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                        SubCategoryName = "Booking Sheet",
                                        EventDescription = $"\nProgress note not created yet or Progress date {ProgressNotesDates[j]} is greated than current date.",
                                    });
                                    Log.Information($"\nProgress note not created yet or Progress date {ProgressNotesDates[j]} is greated than current date.");
                                    it++;
                                    if (it < patientsXpaths.Count)
                                    {
                                        goto oneprogressnote;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                if (ProgressNotesDates[j] != "Telephone/Web" && DateTime.Compare(Convert.ToDateTime(ProgressNotesDates[j]), Convert.ToDateTime(date)) < 0)
                                {
                                    if (DateTime.Now.Date < Convert.ToDateTime(ProgressNotesDates[j]))
                                    {
                                        j++;
                                        continue;
                                    }
                                    progressNoteIteratorValue = j + 1;
                                    break;
                                }
                                j++;
                                continue;
                            }
                        retry:
                            var AllergiesList = new List<string>();
                            var Allergies = "";
                            try
                            {
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = "Retrieving Allergies",
                                });
                                Log.Information("Retrieving Allergies");
                                for (int i = 1; i < 10; i++)
                                {
                                    string? allergy = Convert.ToString(driver.ExecuteJavaScript<object>($"return document.querySelector('#overview_rpTbl18 > tbody > tr:nth-child({i}) > td:nth-child(2) > div').title"));
                                    AllergiesList.Add(allergy);
                                }
                            }
                            catch
                            {
                                goto AllergyResume;
                            }
                        AllergyResume:
                            if (AllergiesList.Count > 0)
                            {
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = AllergiesList.Count + " Allergies Found",
                                });
                                Log.Information(AllergiesList.Count + " Allergies Found");
                                Allergies = string.Join(",", AllergiesList);
                            }
                            if (AllergiesList.Count < 1)
                            {
                                Log.Information("No Allergies found");
                            }
                            try
                            {
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = "Retrieving CPT Codes and Description",
                                });
                                Log.Information("Retrieving CPT Codes and Description");
                                driver.FindElement(By.XPath($"//*[@id='encDropDownList']/li[{progressNoteIteratorValue}]")).Click();
                                List<string> CPT = new List<string>();
                                var trycpt = "//*[contains(text(), 'PROCEDURE:')]";
                                try
                                {
                                    driver.FindElement(By.XPath("/html/body/div[20]/div/div/div[3]/button[2]")).Click();
                                }
                                catch (Exception)
                                {
                                    Log.Information("Exception Handled for JS alert");
                                    goto cpp;
                                }
                            cpp:

                                Thread.Sleep(5000);
                                var CPTsListElements = driver.FindElements(By.XPath(trycpt));
                                if (CPTsListElements.Count == 0)
                                {
                                    CPTsListElements = driver.FindElements(By.XPath("//*[contains(text(), 'Procedure:')]"));
                                }

                                List<string> uniqueCPTValues = new List<string>();

                                foreach (var element in CPTsListElements)
                                {
                                    Thread.Sleep(500);
                                    string text = element.Text.Trim();
                                    // Check if the current text value is already present in the uniqueTextValues list
                                    if (!uniqueCPTValues.Contains(text.Trim()))
                                    {
                                        if (!(text.Trim().Contains('/')))
                                        {
                                            uniqueCPTValues.Add(text.Trim());
                                        }
                                    }
                                }

                                for (int ele = 0; ele < uniqueCPTValues.Count; ele++)
                                {
                                    var Cpts = uniqueCPTValues[ele].Split(":");
                                    if (Cpts[Cpts.Length - 1].Contains(";"))
                                    {
                                        var temp = Cpts[Cpts.Length - 1].Split(';');
                                        Thread.Sleep(500);

                                        foreach (var item in temp)
                                        {
                                            if (item.Trim().ToUpper() == "COLOREC CNCR SCR" || item.Trim().ToUpper() == "COLOREC CANCR SCR")
                                            {
                                                continue;
                                            }
                                            if (CPT.Contains(item.Trim()))
                                            {
                                                continue;
                                            }
                                            CPT.Add(item.Trim());
                                        }
                                        continue;
                                    }
                                    if (Cpts[Cpts.Length - 1].Contains(","))
                                    {
                                        var temp = Cpts[Cpts.Length - 1].Split(',');
                                        Thread.Sleep(500);

                                        foreach (var item in temp)
                                        {
                                            if (item.Trim().ToUpper() == "COLOREC CNCR SCR" || item.Trim().ToUpper() == "COLOREC CANCR SCR")
                                            {
                                                continue;
                                            }
                                            if (CPT.Contains(item.Trim()))
                                            {
                                                continue;
                                            }
                                            CPT.Add(item.Trim());
                                        }
                                        continue;
                                    }
                                    else
                                    {
                                        if (Cpts[Cpts.Length - 1].Trim().ToUpper() != "COLOREC CNCR SCR" || Cpts[Cpts.Length - 1].Trim().ToUpper() != "COLOREC CANCR SCR")
                                        {
                                            CPT.Add(Cpts[Cpts.Length - 1].Trim());

                                        }
                                    }
                                }
                                CPT = CPT.Distinct().ToList();


                                if (Reason.ToLower().Contains("colonoscopy") || Reason.ToLower().Contains("colon") || Reason.ToLower().Contains("egd"))
                                {
                                    if (Reason.Trim().ToLower() == "egd" && CPT.Count == 2 && CPT[0].Trim().ToUpper() == "UPPER GI ENDOSCOPY" && CPT[1].Trim().ToUpper() == "BIOPSY")
                                    {
                                        ;
                                    }
                                    else
                                    {
                                        CPT = FilteredCPTDescription(CPT, Reason);
                                    }
                                }

                                List<string> CPTCodes = new List<string>();
                                for (int cp = 0; cp < CPT.Count; cp++)
                                {
                                    if (CPT[cp].ToString().Trim().ToUpper() == "UPPER GI ENDOSCOPY" && CPT[cp + 1].ToString().Trim().ToUpper() == "BIOPSY")
                                    {
                                        CPTCodes.Add(PBSDataService.getCPTInfo(CPT[cp].Trim() + ", " + CPT[cp + 1].Trim()));
                                        continue;
                                    }
                                    if (CPT[cp].ToString().Trim().ToUpper() == "ERCP" && CPT[cp + 1].ToString().Trim().ToUpper() == "WITH STENT REMOVAL")
                                    {
                                        CPTCodes.Add(PBSDataService.getCPTInfo(CPT[cp] + ", " + CPT[cp + 1]));
                                        continue;
                                    }
                                    else
                                    {
                                        if (CPT[cp].Trim().ToUpper() == "BIOPSY")
                                        {
                                            continue;
                                        }
                                        CPTCodes.Add(PBSDataService.getCPTInfo($"{CPT[cp].ToString().Trim()}"));
                                    }
                                }

                                var CPTCode = string.Join(',', CPTCodes);
                                var CPTDesc = string.Join(',', CPT);
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = "Retrieving ICD Codes and Description",
                                });
                                Log.Information("Retrieving ICD Codes and Description");
                                var ICDCOde = "";
                                var ICDDesc = "";

                                try
                                {
                                    var ic = driver.FindElement(By.XPath("//*[contains(text(), '(Primary)')]"));
                                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", ic);
                                    var icdRaw = driver.FindElement(By.XPath("//*[contains(text(), '(Primary)')]")).Text.ToString();
                                    ICDCOde = icdRaw.Split(" - ")[1].Split("(Primary)")[0].Trim();
                                    ICDDesc = icdRaw.Split(" - ")[0].Split('.')[0].Trim();
                                    if (ICDDesc == "1")
                                    {
                                        ICDDesc = icdRaw.Split(" - ")[0].Split('.')[1].Trim();
                                    }
                                    if (ICDDesc.Contains("Assessment"))
                                    {
                                        ICDDesc = icdRaw.Split(" - ")[0].Split("(Primary)")[0].Trim().Split("1.")[1].Trim();
                                    }

                                    ICDDesc = Regex.Replace(ICDDesc, @"'", "''");
                                }
                                catch (Exception)
                                {
                                    PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                    {
                                        GeneralLogID = gLogId,
                                        EventTime = DateTime.Now,
                                        ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                        CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                        SubCategoryName = "Booking Sheet",
                                        EventDescription = "ICD Primary not Found. Skipping ICD Section",
                                    });
                                    Log.Information("ICD Primary not Found. Skipping ICD Section");
                                    ICDDesc = "(Primary) ICD not found on progress notes";
                                    ICDCOde = "N/A";
                                    goto icdskip;
                                }
                            icdskip:
                                try
                                {
                                    driver.Navigate().Refresh();
                                    driver.SwitchTo().Alert().Accept();
                                }
                                catch (Exception)
                                {
                                    goto conttt;
                                }
                            conttt:
                                Thread.Sleep(5000);
                                driver.FindElement(By.XPath("//*[@id='pat-details']/div/div[1]/div[1]/div[2]/span[1]")).Click(); //info click
                                Thread.Sleep(15000);
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = "Retrieving Patient Information",
                                });
                                Log.Information("Retrieving Patient Information");

                            retrydemographic:
                                List<PatientInfo> obj = new List<PatientInfo>();
                                var accNumber = WaitForElementValue(driver, "//*[@id='patient-demographicsIpt5']");
                                var LastName = WaitForElementValue(driver, "//*[@id='ptlname']");
                                var FirstName = WaitForElementValue(driver, "//*[@id='ptfname']");
                                var Gender = WaitForElementValue(driver, "//*[@id='ptsex']");
                                var DoBstr = WaitForElementValue(driver, "//*[@id='dateofbirth']");
                                var DoB = DoBstr.Replace("/", "-");
                                var PhoneHome = WaitForElementValue(driver, "//*[@id='pthomephone']");
                                var Cellno = WaitForElementValue(driver, "//*[@id='ptcellphone']");
                                var Address1 = WaitForElementValue(driver, "//*[@id='ptaddress']");
                                var Address2 = WaitForElementValue(driver, "//*[@id='ptaddress2']");
                                var Email = WaitForElementValue(driver, "//*[@id='ptemail']");
                                var State = WaitForElementValue(driver, "//*[@id='patient-demographicsIpt8']");
                                var City = WaitForElementValue(driver, "//*[@id='ptcity']");
                                var Zip = WaitForElementValue(driver, "//*[@id='ptzip']");

                                //assinging retrieved data to patient object
                                PatientInfo patobj = new PatientInfo
                                {
                                    PracticeCode = "DDA",
                                    AccountNumber = accNumber,
                                    PatientFirstName = FirstName,
                                    PatientLastName = LastName,
                                    PatientAddressLine1 = Address1,
                                    PatientAddressLine2 = Address2,
                                    PatientCity = City,
                                    PatientState = State,
                                    PatientZip = Zip,
                                    PatientPhone = PhoneHome,
                                    PatientGender = Gender,
                                    PatientDOB = DoB,
                                    PatientEmail = Email,
                                    PatientCellPhone = Cellno,
                                    ProviderName = ProviderName,
                                    ProcedureTime = procedureTime,
                                    ProcedureDate = procedureDate,
                                    ProcedureLength = "30 Minutes",
                                    Allergies = Allergies,
                                    ICDs = ICDCOde,
                                    CPTs = CPTCode,
                                    Diagnoses = ICDDesc,
                                    AnesthesiaType = "",
                                    Diabetic = "",
                                    ProcedureDescription = CPTDesc,
                                    VisitType = visitType,
                                    Reason = Reason,
                                    FacilityName = FacilityName
                                };

                                if (patobj.PatientFirstName.Trim() == "" || patobj.PatientLastName.Trim() == "")
                                {
                                    goto retrydemographic;
                                }

                                if (FacilityName == "Advent Health Wesley Chapel")
                                {
                                    patobj.AnesthesiaType = "MAC";
                                }

                                try
                                {
                                    int primary = 0;
                                    Thread.Sleep(100);
                                    var PrimaryInsuranceName = driver.FindElement(By.XPath("//*[@id='patient-demographicsTbl3']/tbody/tr/td[5]")).Text.ToString();
                                    Thread.Sleep(100);
                                    var PrimaryInsuranceNumber = driver.FindElement(By.XPath("//*[@id='patient-demographicsTbl3']/tbody/tr/td[7]")).Text.ToString();
                                    Thread.Sleep(100);
                                    var PrimaryGroupNumber = driver.FindElement(By.XPath("//*[@id=\"patient-demographicsTbl3\"]/tbody/tr/td[11]")).Text.ToString();
                                    Thread.Sleep(100);
                                    var isPrimary = driver.FindElement(By.XPath("//*[@id='patient-demographicsTbl3']/tbody/tr/td[4]")).Text.ToString();
                                    primary = 1;

                                    if (isPrimary.ToLower() == "p" && isPrimary != null && primary == 1)
                                    {
                                        string alphanumberic = "[^a-zA-Z0-9]";
                                        patobj.PrimaryInsuaranceName = PrimaryInsuranceName;
                                        patobj.PrimaryInsuaranceNumber = PrimaryInsuranceNumber;
                                        patobj.PrimaryGroupNumber = PrimaryGroupNumber;
                                        patobj.PrimaryInsuaranceNumber = Regex.Replace(patobj.PrimaryInsuaranceNumber, alphanumberic, "");
                                    }

                                    int secondary = 0;
                                    Thread.Sleep(100);
                                    var SecondaryInsuranceName = driver.FindElement(By.XPath("//*[@id='patient-demographicsTbl3']/tbody/tr[2]/td[5]")).Text.ToString();
                                    Thread.Sleep(100);
                                    var SecondaryInsuranceNumber = driver.FindElement(By.XPath("//*[@id='patient-demographicsTbl3']/tbody/tr[2]/td[7]")).Text.ToString();
                                    Thread.Sleep(100);
                                    var isSecondary = driver.FindElement(By.XPath("//*[@id='patient-demographicsTbl3']/tbody/tr[2]/td[4]")).Text.ToString();
                                    Thread.Sleep(100);
                                    var SecondaryGroupNumber = driver.FindElement(By.XPath("//*[@id=\"patient-demographicsTbl3\"]/tbody/tr[2]/td[11]")).Text.ToString();
                                    secondary = 1;

                                    if (isSecondary.ToLower() == "s" && isSecondary != null && secondary == 1)
                                    {
                                        string alphanumberic = "[^a-zA-Z0-9]";
                                        patobj.SecondaryInsuaranceName = SecondaryInsuranceName;
                                        patobj.SecondaryInsuaranceNumber = SecondaryInsuranceNumber;
                                        patobj.SecondaryGroupNumber = PrimaryGroupNumber;
                                        patobj.SecondaryInsuaranceNumber = Regex.Replace(patobj.SecondaryInsuaranceNumber, alphanumberic, "");
                                    }
                                    obj.Add(patobj);
                                    PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                    {
                                        GeneralLogID = gLogId,
                                        EventTime = DateTime.Now,
                                        ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                        CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                        SubCategoryName = "Booking Sheet",
                                        EventDescription = "Retrieving Progress note PDF",
                                    });
                                    Log.Information("Retrieving Progress note PDF");
                                    ProgreesNotePDF(driver, patobj);
                                    PBSDataService.SavePatientInfo("RPA", obj);
                                    PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                    {
                                        GeneralLogID = gLogId,
                                        EventTime = DateTime.Now,
                                        ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                        CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                        SubCategoryName = "Booking Sheet",
                                        EventDescription = $"Data of Patient # {it + 1} out of {patientsXpaths.Count} Saved to DB | Name : {patobj.PatientLastName} , {patobj.PatientFirstName} | Provider = {ProviderName}  Facility = {FacilityName}",
                                    });
                                    Log.Information($"Data of Patient # {it + 1} out of {patientsXpaths.Count} Saved to DB | Name : {patobj.PatientLastName} , {patobj.PatientFirstName} | Provider = {ProviderName}  Facility = {FacilityName}");
                                }
                                catch (NoSuchElementException)
                                {
                                    obj.Add(patobj);
                                    PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                    {
                                        GeneralLogID = gLogId,
                                        EventTime = DateTime.Now,
                                        ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                        CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                        SubCategoryName = "Booking Sheet",
                                        EventDescription = "Retrieving Progress note PDF",
                                    });
                                    Log.Information("Retrieving Progress note PDF");
                                    ProgreesNotePDF(driver, patobj);
                                    PBSDataService.SavePatientInfo("RPA", obj);
                                    PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                    {
                                        GeneralLogID = gLogId,
                                        EventTime = DateTime.Now,
                                        ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                        CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                        SubCategoryName = "Booking Sheet",
                                        EventDescription = $"Data of Patient # {it + 1} out of {patientsXpaths.Count} Saved to DB | Name : {patobj.PatientLastName} , {patobj.PatientFirstName} | Provider = {ProviderName}  Facility = {FacilityName}",
                                    });
                                    Log.Information($"Data of Patient # {it + 1} out of {patientsXpaths.Count} Saved to DB | Name : {patobj.PatientLastName} , {patobj.PatientFirstName} | Provider = {ProviderName} | Facility = {FacilityName}");
                                    goto ss;
                                }
                            ss:
                                ;
                            }
                            catch (NoSuchElementException ex)
                            {
                                if (count < 3)
                                {
                                    count++;
                                    goto retry;
                                }
                                gLogId = PBSDataService.SaveGeneralLog(new GeneralLog
                                {
                                    DataListMappingID = 110,
                                    StartTime = StartTime,
                                    Status = "Exception Occured",
                                    EndTime = DateTime.Now
                                });
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}",
                                });
                                Log.Information($"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                                continue;
                            }
                            catch (ElementNotInteractableException ex)
                            {
                                if (count < 3)
                                {
                                    count++;
                                    goto retry;
                                }
                                gLogId = PBSDataService.SaveGeneralLog(new GeneralLog
                                {
                                    DataListMappingID = 110,
                                    StartTime = StartTime,
                                    Status = "Exception Occured",
                                    EndTime = DateTime.Now
                                });
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}",
                                });
                                Log.Information($"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                                continue;
                            }
                            catch (ElementNotSelectableException ex)
                            {
                                if (count < 3)
                                {
                                    count++;
                                    goto retry;
                                }
                                gLogId = PBSDataService.SaveGeneralLog(new GeneralLog
                                {
                                    DataListMappingID = 110,
                                    StartTime = StartTime,
                                    Status = "Exception Occured",
                                    EndTime = DateTime.Now
                                });
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}",
                                });
                                Log.Information($"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                                continue;
                            }
                            catch (UnhandledAlertException ex)
                            {
                                if (count < 3)
                                {
                                    count++;
                                    goto retry;
                                }
                                gLogId = PBSDataService.SaveGeneralLog(new GeneralLog
                                {
                                    DataListMappingID = 110,
                                    StartTime = StartTime,
                                    Status = "Exception Occured",
                                    EndTime = DateTime.Now
                                });
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}",
                                });
                                Log.Information($"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                                continue;
                            }
                            catch (NullReferenceException ex)
                            {
                                if (count < 3)
                                {
                                    count++;
                                    goto retry;
                                }
                                gLogId = PBSDataService.SaveGeneralLog(new GeneralLog
                                {
                                    DataListMappingID = 110,
                                    StartTime = StartTime,
                                    Status = "Exception Occured",
                                    EndTime = DateTime.Now
                                });
                                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                                {
                                    GeneralLogID = gLogId,
                                    EventTime = DateTime.Now,
                                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                                    SubCategoryName = "Booking Sheet",
                                    EventDescription = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}",
                                });
                                Log.Information($"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                                continue;
                            }
                            finally
                            {
                                driver.Navigate().Refresh();
                                try
                                {
                                    driver.SwitchTo().Alert().Accept();
                                }
                                catch
                                {
                                    goto final;
                                }
                            final:
                                driver.Navigate().GoToUrl(URL);
                                Thread.Sleep(5000);
                                js.ExecuteScript("window.onbeforeunload = function() {};");

                            }
                        }
                    }
                }
                Success = 1;
            }

            catch (ElementNotInteractableException ex)
            {
                gLogId = PBSDataService.SaveGeneralLog(new GeneralLog
                {
                    DataListMappingID = 110,
                    StartTime = StartTime,
                    Status = "Exception Occured",
                    EndTime = DateTime.Now
                });
                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                {
                    GeneralLogID = gLogId,
                    EventTime = DateTime.Now,
                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                    SubCategoryName = "Booking Sheet",
                    EventDescription = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}",
                });
                Log.Information($"Exception Occured due to Element not interactable at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                Success = 0;
            }
            catch (ElementNotSelectableException ex)
            {
                gLogId = PBSDataService.SaveGeneralLog(new GeneralLog
                {
                    DataListMappingID = 110,
                    StartTime = StartTime,
                    Status = "Exception Occured",
                    EndTime = DateTime.Now
                });
                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                {
                    GeneralLogID = gLogId,
                    EventTime = DateTime.Now,
                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                    SubCategoryName = "Booking Sheet",
                    EventDescription = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}",
                });
                Log.Information($"Exception Occured due to Element not interactable at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                Success = 0;
            }
            catch (IndexOutOfRangeException ex)
            {
                gLogId = PBSDataService.SaveGeneralLog(new GeneralLog
                {
                    DataListMappingID = 110,
                    StartTime = StartTime,
                    Status = "Exception Occured",
                    EndTime = DateTime.Now
                });

                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                {
                    GeneralLogID = gLogId,
                    EventTime = DateTime.Now,
                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                    SubCategoryName = "Booking Sheet",
                    EventDescription = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}",
                });
                Log.Information($"Exception Occured due to Index out of range at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                Success = 0;
            }
            catch (FormatException ex)
            {
                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                {
                    GeneralLogID = gLogId,
                    EventTime = DateTime.Now,
                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                    SubCategoryName = "Booking Sheet",
                    EventDescription = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}",
                });
                Log.Information($"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                Success = 0;
            }
            catch (Exception ex)
            {
                gLogId = PBSDataService.SaveGeneralLog(new GeneralLog
                {
                    DataListMappingID = 110,
                    StartTime = StartTime,
                    Status = "Exception Occured",
                    EndTime = DateTime.Now
                });

                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                {
                    GeneralLogID = gLogId,
                    EventTime = DateTime.Now,
                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                    SubCategoryName = "Booking Sheet",
                    EventDescription = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}",
                });
                Log.Information($"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}");
                PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                {
                    GeneralLogID = gLogId,
                    EventTime = DateTime.Now,
                    ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                    CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                    SubCategoryName = "Booking Sheet",
                    EventDescription = "Restarting application",
                });
                Log.Information("\n\n\nRestarting application.");
                PBSDataService dsObj = new();
                var emailmessage = $"Exception Occured at Line {Convert.ToInt32(ex?.StackTrace?.Substring(ex.StackTrace.LastIndexOf(' ')))}: {ex?.Message}\n\n Application Restarting now.";
                var emailsubject = "Exception Occured : Booking Sheet DDA";
                dsObj.EmailNotification(emailmessage, emailsubject);
                Status(date);
            }
            finally
            {
                driver.Quit();
            }
            return Success;
        }

        //Filter CPTs based on reason (mentioned in appSettings.json)
        static List<string> FilteredCPTDescription(List<string> CPT, string Reason)
        {
            List<string> FilteredCPTList = new List<string>();

            if (Reason.ToLower().Contains("colonoscopy") || Reason.ToLower().Contains("colon"))
            {
                List<string> ListFromDB = new List<string>();
                string[]? CPTCodes = PBSDataService.GetCPTfromAppSettings("colon").Split(",");
                foreach (var item in CPTCodes)
                {
                    var onetimelist = PBSDataService.CPTInfoFiltered(item);
                    foreach (var element in onetimelist)
                    {
                        ListFromDB.Add(element.CPTDescription);
                    }
                }
                foreach (string item in CPT)
                {
                    if (ListFromDB.Contains(item))
                    {
                        FilteredCPTList.Add(item);
                    }
                }
            }
            if (Reason.ToLower().Contains("egd"))
            {
                FilteredCPTList.Add("UPPER GI ENDOSCOPY, BIOPSY");
            }
            if (FilteredCPTList.Count == 0)
            {
                FilteredCPTList.Add("CPT not found on Progress note");
            }
            return FilteredCPTList;
        }

        //Check if URL is working or not
        static async Task<bool> CheckUrl()
        {
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    HttpResponseMessage response = await client.GetAsync(URL);
                    return response.StatusCode == HttpStatusCode.OK;
                }
                catch (HttpRequestException)
                {
                    return false;
                }
            }
        }

        //Retrieve Patient Progress Note ( patient object is used to modify
        static void ProgreesNotePDF(IWebDriver driver, PatientInfo obj)
        {
            string baseDirectory = Path.Combine(Environment.CurrentDirectory, "Facility HTMLs\\Patient ProgressNotes PDFs");
            string facilityName = obj.FacilityName;
            string procedureDate = obj?.ProcedureDate?.Replace("/", "_");
            string lastName = obj?.PatientLastName;
            string firstName = obj?.PatientFirstName;

            try
            {
                driver.FindElement(By.Id("patient-demographicsBtn57")).Click();
                Thread.Sleep(500);
                string subDirectory = Path.Combine(baseDirectory, $"{facilityName} {procedureDate}");
                if (!Directory.Exists(subDirectory))
                {
                    Directory.CreateDirectory(subDirectory);
                }

                string filenameBase = $"{lastName}, {firstName}";
                int number = GetNextFileNumber(subDirectory);
                string newFileName = Path.Combine(subDirectory, $"{number}. {filenameBase}.pdf");
                var html = "";

            retryPNPDF:
                try
                {
                    if (Convert.ToBoolean(driver.ExecuteJavaScript<object>($"return document.querySelector('#stickyPatientDetails') != null")) == true)
                    {
                        //remove date and name bar (NOT Required)
                        driver.ExecuteJavaScript("document.querySelector('#stickyPatientDetails').setHTML('')");
                    }
                    //normal pdf page
                    html = driver.FindElement(By.XPath("//*[@id='progress_content']/div/table/tbody/tr/td")).GetAttribute("innerHTML");
                }
                catch
                {
                    //in case normal pdf page is not retreived
                    html = driver.FindElement(By.XPath("//*[@id='progress_content']/div")).GetAttribute("innerHTML");
                    goto GeneratePDF;
                }
            //HTML to PDF Conversion using NReco PDF Generator
            GeneratePDF:
                if (html != null || html?.Trim() != "")
                {
                    var pdfGenerator = new HtmlToPdfConverter();
                    // Set the page size and orientation
                    pdfGenerator.Orientation = PageOrientation.Portrait;
                    pdfGenerator.Size = PageSize.A4;
                    // Convert the HTML string to PDF
                    var pdfBytes = pdfGenerator.GeneratePdf(html);
                    // Save the PDF to a file
                    File.WriteAllBytes(newFileName, pdfBytes);
                    PBSDataService.SaveGeneralLogDetails(new GeneralLogDetail
                    {
                        GeneralLogID = gLogId,
                        EventTime = DateTime.Now,
                        ProcessName = "Automation Process for : " + "Patient Booking Sheet",
                        CategoryName = "PBS" + " | Practice ID : 6 | " + "DDA",
                        SubCategoryName = "Booking Sheet",
                        EventDescription = $"Generated Progress Notes PDF for {obj?.PatientLastName}, {obj?.PatientFirstName}",
                    });
                    Log.Information($"Generated Progress Notes PDF for {obj?.PatientLastName}, {obj?.PatientFirstName}");
                    Log.Information($"File Saved : {Path.GetFileName(newFileName)}");

                }
                else
                {
                    goto retryPNPDF;
                }
            }
            catch (Exception ex)
            {
                Log.Information("Exception during genertating PDF progress note :\t" + ex.Message);
            }
        }
        //Demographic values sometimes being retrieved null (this fucntion fixes it)
        private static string WaitForElementValue(IWebDriver driver, string xPathExpression)
        {
            int count = 5;
        //retries 5 times
        elementretry:
            string? val = new WebDriverWait(driver, TimeSpan.FromSeconds(5)).Until(c => c.FindElement(By.XPath(xPathExpression))).GetAttribute("value");
            if (val != null || val?.Trim() != "")
            {
                return val;
            }
            else
            {
                if (count < 1)
                {
                    count--;
                    Thread.Sleep(100);
                    goto elementretry;
                }
                return "";
            }
        }

        //File Number for
        public static int GetNextFileNumber(string directory)
        {
            int nextNumber = 1;

            if (Directory.Exists(directory))
            {
                string?[] existingFiles = Directory.GetFiles(directory, "*.pdf")
                                                  .Select(Path.GetFileNameWithoutExtension)
                                                  .Where(name => int.TryParse(name.Split('.').First(), out _))
                                                  .ToArray();

                if (existingFiles.Length > 0)
                {
                    int maxNumber = existingFiles
                        .Select(name => int.Parse(name.Split('.').First()))
                        .Max();
                    nextNumber = maxNumber + 1;
                }
            }
            return nextNumber;
        }
    }
}