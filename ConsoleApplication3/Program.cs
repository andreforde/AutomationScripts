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
using System.Runtime.InteropServices;

namespace ConsoleApplication3
{
    class Program
    {
        //Here is the once-per-class call to initialize the log object
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public static object MessageBox { get; private set; }
        public static void ErrorLogging(Exception ex)
        {
            string strPath = "Log.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("===========End============= " + DateTime.Now);
            }
            Emailing.Email.SendEmail("lightbot@lightsourcehr.com", "mehta2626@gmail.com", "dddds", "dsdsssd", strPath);
        }
        public static void ReadError()
        {
            string strPath = "Log.txt";
            using (StreamReader sr = new StreamReader(strPath))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    Console.WriteLine(line);
                }
            }
        }

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

            gc.FindElement(By.XPath("//*[@id='Subject']")).SendKeys("#" + t.Case_no.ToString());

            gc.FindElement(By.XPath("//*[@id='btnSave']")).Click();
        }
        static void Main(string[] args)
        {
           
            //log4net.Config.XmlConfigurator.Configure();
        Wait:
            //while (true)
            //{

            //    String wait = DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString();
            //    if (wait == "2350")
            //    {
            //        break;
            //    }

            //}
            int record = 0;
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            Console.WriteLine("application running at : " + AppDomain.CurrentDomain.BaseDirectory);
            //removing previous data from table
            dal.RemovePreviousOccurences();

            //................
            IWebDriver gc = new ChromeDriver();
            //...loging in 
            gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/Login");
            gc.FindElement(By.Name("LoginID")).SendKeys("lightbot");
            gc.FindElement(By.Name("Password")).SendKeys("RPAuser!");
            gc.FindElement(By.Name("Password")).SendKeys(Keys.Enter);
            //.............
            gc.Navigate().GoToUrl("https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\\AEE1+Ancillary+Risk+Fees");
            gc.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            gc.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            //  gc.FindElement(By.Id("ndbfc0")).Click();
            //    gc.FindElements(By.TagName("option"))[3].Click();
            //      gc.FindElement(By.Id("updateBtnP")).Click();
            //Console.WriteLine("going to sleep");
            System.Threading.Thread.Sleep(2000);
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
            //.........
            //var db1 = new Entities1();
            var delimiter = "\t";

            Console.WriteLine("Data extraction done ;");
            Console.WriteLine("Data with location :");
            foreach (Table i in dal.FetchTableDataWhereLocIsNotNull())
            {
                record++;
                Console.WriteLine(i.billDate + delimiter + i.eventCode + delimiter + i.billRates + delimiter + i.billUnits + delimiter + i.clientID + delimiter + i.location);

            }
            Console.WriteLine("Data without location :");
            foreach (Table i in dal.FetchTableDataWhereLocIsNull())
            {
                Console.WriteLine(i.billDate + delimiter + i.eventCode + delimiter + i.billRates + delimiter + i.billUnits + delimiter + i.clientID + delimiter + i.location);

            }
            foreach (Table i in dal.FetchTableDataWhereLocIsNull())
            {
                //gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next");
                //System.Threading.Thread.Sleep(5000);
                //gc.FindElements(By.ClassName("cs-header-module-item"))[2].Click();
                Program test = new Program();
                Exception e = new Exception("Client Data without location! \n Client ID:" + i.clientID + " Case no :" + i.Case_no);
                ErrorLogging(e);
                test.Create(gc, i);
                /*
                 * code for sending mail for an erro , configure it for sending it out
                try
                {
                    MailMessage mail = new MailMessage("you@yourcompany.com", "xeeshan.ah@gmail.com");
                    SmtpClient client = new SmtpClient("Deadpool.lightsourcemanagement.com");

                    client.Port = 25;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Host = "lightsource.com";
                    mail.Subject = "this is a test email.";
                    mail.Body = "this is my test email body";
                    client.Send(mail);
                }
                catch (Exception ex)
                {

                    Console.WriteLine("err");
                }*/
            }
            //saving file.
            using (var writer = new System.IO.StreamWriter(basePath + "\\data.txt"))
            {
                foreach (Table i in dal.FetchTableDataWhereLocIsNotNull())
                {
                    writer.WriteLine(i.billDate + delimiter + i.eventCode + delimiter + i.billRates + delimiter + i.billUnits + delimiter + i.clientID + delimiter + delimiter + i.location);
                }
            }
            gc.Close();
            gc.Dispose();
            if (record == 0)
            {
                Console.WriteLine("no data found");
                goto end;
            }

        Start:
            //part to connection to cannary chrome
            //...
            //var chromeOptions = new ChromeOptions
            //{
            //    BinaryLocation = @"Canary\chrome.exe"
            //};

            //chromeOptions.AddArguments(new List<string>() { "headless", "disable-gpu" });

            //ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            //service.HideCommandPromptWindow = true;

            //IWebDriver gc1 = new ChromeDriver(service, chromeOptions);
            //........
            IWebDriver gc1 = new ChromeDriver();
            gc1.Navigate().GoToUrl("https://ctw.prismhr.com/ctw/dbweb.asp?dbcgm=1");
            System.Threading.Thread.Sleep(1000);
            gc1.FindElement(By.XPath("//*[@id='text4v1']")).SendKeys("lightbot");
            gc1.FindElement(By.XPath("//*[@id='password6v1']")).SendKeys("RPAuser1!");
            gc1.FindElement(By.XPath("//*[@id='button8v1']")).Click();
            System.Threading.Thread.Sleep(1000);

            gc1.FindElement(By.XPath("//*[@id='text31v1']")).Click();
            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("C");
            gc1.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);

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
            System.Threading.Thread.Sleep(1000);

            gc1.FindElement(By.XPath("//*[@id='button31v2']")).Click();

            System.Threading.Thread.Sleep(1000);
            Console.WriteLine(gc1.Url);
            String current = gc1.CurrentWindowHandle;
            foreach (String winHandle in gc1.WindowHandles)
            {
                gc1.SwitchTo().Window(winHandle);
            }
            //sometimes the upload window doesn't open
            if (gc1.CurrentWindowHandle != current)
            {
                gc1.FindElement(By.XPath("//*[@id='fname']")).SendKeys(basePath + "\\data.txt"); //put the full path of file here

                gc1.FindElement(By.XPath("//*[@id='submit1']")).Click();

                System.Threading.Thread.Sleep(1000);
                gc1.FindElement(By.XPath("//*[@id='BUTTON1']")).Click();
                System.Threading.Thread.Sleep(20000);
                gc1.SwitchTo().Window(current);
                try
                {
                    Exception s = new Exception(gc1.FindElement(By.XPath("//*[@id='body_span29v2']")).Text);
                    ErrorLogging(s);
                    //send this exception in the mail
                    /*
                    try
                    {
                        MailMessage mail = new MailMessage("you@yourcompany.com", "xeeshan.ah@gmail.com");
                        SmtpClient client = new SmtpClient("Deadpool.lightsourcemanagement.com");

                        client.Port = 25;
                        client.DeliveryMethod = SmtpDeliveryMethod.Network;
                        client.UseDefaultCredentials = false;
                        client.Host = "lightsource.com";
                        mail.Subject = "this is a test email.";
                        mail.Body = "this is my test email body";
                        client.Send(mail);
                    }
                    catch (Exception ex)
                    {

                        Console.WriteLine("err");
                    }
                    */
                }
                catch (Exception e)
                {

                }
                try
                {
                    gc1.FindElement(By.XPath("//*[@id='button33v2']")).Click();
                    System.Threading.Thread.Sleep(2000);
                    gc1.FindElement(By.XPath("//*[@id='button32v2']")).Click();
                    System.Threading.Thread.Sleep(2000);
                    gc1.SwitchTo().Alert().Accept();
                    gc1.SwitchTo().Window(current);
                    gc1.FindElement(By.XPath("//*[@id='button35v2']")).Click();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    gc1.SwitchTo().Window(current);
                    gc1.FindElement(By.XPath("//*[@id='button35v2']")).Click();

                }
                gc1.Close();
                gc1.Dispose();
                Console.WriteLine("Process Complete..!");


            }
            else
            {
                gc1.Close();
                gc1.Dispose();
                goto Start;
            }

        end:
            System.Threading.Thread.Sleep(40000);

            goto Wait;
        }

    }
}
