using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using DAL;
using System.IO;

namespace ConsoleApplication3
{
    class Program
    {
        //Here is the once-per-class call to initialize the log object
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public void Create(IWebDriver gc, Table t)
        {
            gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/peo/client");
            gc.FindElement(By.Id("dropdownMenu1")).Click();
            gc.FindElement(By.Id("All")).Click();
            gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
            gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
            gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
            gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
            gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
            gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
            gc.FindElement(By.Name("ClientNumber")).SendKeys(t.clientID.ToString()); // put the client number here.
            System.Threading.Thread.Sleep(2000);
            gc.FindElement(By.ClassName("formSearchBtn")).Click();
            System.Threading.Thread.Sleep(1000);
            gc.FindElements(By.ClassName("cs-underline"))[1].Click();
            gc.FindElement(By.XPath("//*[@id='customHeaderHtml']/div[2]/div[6]/div/div[1]/table/tbody/tr/td[1]/span")).Click();
            gc.FindElement(By.XPath("//*[@id='lstDataform_btnAdd']")).Click();
            gc.FindElement(By.XPath("//*[@id='crCategory']")).SendKeys("R");
            System.Threading.Thread.Sleep(1500);
            gc.FindElement(By.XPath("//*[@id='crCategory']")).SendKeys(Keys.Tab);
            gc.FindElement(By.XPath("//*[@id='fkCaseTypeID']")).SendKeys("M");
            System.Threading.Thread.Sleep(1500);
            gc.FindElement(By.XPath("//*[@id='fkCaseTypeID']")).SendKeys(Keys.Enter);
            gc.FindElement(By.XPath("//*[@id='CallerName']")).SendKeys("Lightbot");
            gc.FindElement(By.XPath("//*[@id='EmailAddress']")).SendKeys("lightbot@lightsourcehr.com");
            DateTime dateTime = DateTime.Today;
            gc.FindElement(By.XPath("//*[@id='DueDate']")).SendKeys(dateTime.ToString("MM/dd/yyyy"));
            gc.FindElement(By.XPath("//*[@id='btnSave']")).Click();
        }
        static void Main(string[] args)
        {
            log4net.Config.XmlConfigurator.Configure();
           
            try
            {
                //removing previous data from table
                DAL.dal.RemovePreviousOccurences();
            }
            catch(Exception ex)
            {
                log.Error(ex.Message);
            }

            //................
            try
            {
                var chromeOptions1 = new ChromeOptions
                {
                    BinaryLocation = @"Canary\chrome.exe"
                };

                chromeOptions1.AddArguments(new List<string>() { "headless", "disable-gpu" });

                ChromeDriverService service1 = ChromeDriverService.CreateDefaultService();
                service1.HideCommandPromptWindow = true;

                IWebDriver gc = new ChromeDriver(service1, chromeOptions1);

                //...loging in 
                gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/Login");
                gc.FindElement(By.Name("LoginID")).SendKeys("lightbot");
                gc.FindElement(By.Name("Password")).SendKeys("RPAuser!");
                gc.FindElement(By.Name("Password")).SendKeys(Keys.Enter);
                //.............
                gc.Navigate().GoToUrl("https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\\AEE1+Ancillary+Risk+Fees");
                gc.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                gc.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
                gc.FindElement(By.Id("ndbfc0")).Click();
                gc.FindElements(By.TagName("option"))[3].Click();
                gc.FindElement(By.Id("updateBtnP")).Click();
                Console.WriteLine("going to sleep");
                System.Threading.Thread.Sleep(10000);
                string check = null;
                ICollection<IWebElement> a = gc.FindElements(By.ClassName("ReportItem"));
                Console.WriteLine("value : " + a.Count);
                foreach (IWebElement b in a)
                {
                    check = b.Text;
                    check = check.Replace(' ', '_');
                    int count = 0;
                    foreach (char c in check)
                    {
                        if (c.Equals('_'))
                        {
                            count++;
                        }
                    }
                    string[] p = check.Split('_');
                    if (count == 5)
                    {
                        var table = new Table();
                        table.Case_no = p[0];
                        table.billDate = p[1];
                        table.eventCode = p[2];
                        table.billRates = p[3];
                        table.billUnits = p[4];
                        table.clientID = p[5];
                        table.location = null;
                        dal.SaveTableData(table);
                    }
                    else
                        if (count == 6)
                    {
                        var table = new Table();
                        table.Case_no = p[0];
                        table.billDate = p[1];
                        table.eventCode = p[2];
                        table.billRates = p[3];
                        table.billUnits = p[4];
                        table.clientID = p[5];
                        table.location = p[6];
                        dal.SaveTableData(table);
                    }
                }
                foreach (IWebElement c in gc.FindElements(By.ClassName("AlternatingItem")))
                {

                    check = c.Text;

                    check = check.Replace(' ', '_');
                    int count = 0;
                    foreach (char d in check)
                    {
                        if (d == '_')
                        {
                            count++;
                        }
                    }
                    string[] p = check.Split('_');
                    if (count == 5)
                    {
                        var table = new Table();
                        table.Case_no = p[0];
                        table.billDate = p[1];
                        table.eventCode = p[2];
                        table.billRates = p[3];
                        table.billUnits = p[4];
                        table.clientID = p[5];
                        table.location = null;
                        dal.SaveTableData(table);
                    }
                    else
                        if (count == 6)
                    {
                        var table = new Table();
                        table.Case_no = p[0];
                        table.billDate = p[1];
                        table.eventCode = p[2];
                        table.billRates = p[3];
                        table.billUnits = p[4];
                        table.clientID = p[5];
                        table.location = p[6];
                        dal.SaveTableData(table);
                    }

                }


            }
            catch(Exception ex)
            {
                log.Error("Error occured while exctracting data:" + ex.Message);
            }

            Console.WriteLine("Data extraction done ;");
            Console.WriteLine("Data with location :");
            //foreach (Table i in db1.Tables.Where(s => s.location == null))
            //{
            //    //gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next");
            //    //System.Threading.Thread.Sleep(5000);
            //    //gc.FindElements(By.ClassName("cs-header-module-item"))[2].Click();
            //    Program test = new Program();
            //    //test.Create(gc,i);
            //}

            //Clear previous failure logs
            string[] filePaths = Directory.GetFiles(@"FailureLogs\");
            foreach (string filePath in filePaths)
                File.Delete(filePath);

            //Create an error log for records with loc = null
            var delimiter = "\t";
            List<Table> lstDataLocNull = new List<Table>();
            string FileName = Guid.NewGuid().ToString();
            try
            {
                lstDataLocNull = dal.FetchTableDataWhereLocIsNull();
                
                using (var writer = new System.IO.StreamWriter("FailureLogs" + @"\" + FileName + ".txt"))
                {
                    foreach (Table i in lstDataLocNull)
                    {
                        writer.WriteLine(i.Case_no + delimiter + i.billDate + delimiter + i.eventCode + delimiter + i.billRates + delimiter + i.billUnits + delimiter + i.clientID + delimiter + i.location);
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("Error occured while creating error records csv:" + ex.Message);
            }

            //Send invalid records log email
            if (Directory.GetFiles("FailureLogs").Count() != 0)
            {
                //Emailing.Email.SendEmail("", "", "", "", "");
            }

            //Create csv for successfull records with loc <> null
            try
            {
                using (var writer = new System.IO.StreamWriter("hello.txt"))
                {
                    List<Table> lstData = dal.FetchTableDataWhereLocIsNotNull();
                    foreach (Table i in lstData)
                    {
                        writer.WriteLine(i.Case_no + delimiter + i.billDate + delimiter + i.eventCode + delimiter + i.billRates + delimiter + i.billUnits + delimiter + i.clientID + delimiter + i.location);
                    }
                }
            }
            catch(Exception ex)
            {
                log.Error("Error occured while creating csv:" + ex.Message);
            }



        Start:
            var chromeOptions = new ChromeOptions
            {
                BinaryLocation = @"Canary\chrome.exe"
            };

            chromeOptions.AddArguments(new List<string>() { "headless", "disable-gpu" });

            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;

            IWebDriver gc1 = new ChromeDriver(service, chromeOptions);
            gc1.Navigate().GoToUrl("https://ctw.prismhr.com/ctw/dbweb.asp?dbcgm=1");
            System.Threading.Thread.Sleep(1000);
            gc1.FindElement(By.XPath("//*[@id='text4v1']")).SendKeys("lightbot");
            gc1.FindElement(By.XPath("//*[@id='password6v1']")).SendKeys("RPAuser1!");
            gc1.FindElement(By.XPath("//*[@id='button8v1']")).Click();
            System.Threading.Thread.Sleep(1000);

            gc1.FindElement(By.XPath("//*[@id='text31v1']")).Click();
            //gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("C");

            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("c");
            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("l");

            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("i");

            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("e");
            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("n");
            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("t");

            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(" bill");
            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);

            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);
            System.Threading.Thread.Sleep(1000);

            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Enter);
            System.Threading.Thread.Sleep(10000);

            //gc1.FindElement(By.XPath("//*[@id='button31v2']"));
            System.Threading.Thread.Sleep(10000);

            gc1.FindElement(By.XPath("//*[@id='button31v2']")).Click();

            System.Threading.Thread.Sleep(10000);
            Console.WriteLine(gc1.Url);
            String current = gc1.CurrentWindowHandle;
            foreach (String winHandle in gc1.WindowHandles)
            {
                gc1.SwitchTo().Window(winHandle);
            }
            //sometimes the upload window doesn't open
            if (gc1.CurrentWindowHandle != current)
            {
                gc1.FindElement(By.XPath("//*[@id='fname']")).SendKeys(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\\hello.txt"); //put the full path of file here

                gc1.FindElement(By.XPath("//*[@id='submit1']")).Click();

                System.Threading.Thread.Sleep(1000);
                gc1.FindElement(By.XPath("//*[@id='BUTTON1']")).Click();

            }
            else
            {
                gc1.Close();
                gc1.Dispose();
                goto Start;
            }
        }
        
    }
}
