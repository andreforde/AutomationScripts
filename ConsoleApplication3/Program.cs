using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;

namespace ConsoleApplication3
{
    class Program
    {
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
            //removing previous data from table
            using (var db = new Entities1())
            {
                List<Table> ta = db.Tables.ToList();
                foreach(Table i in ta)
                {
                    db.Tables.Remove(i);
                }
                db.SaveChanges();
            }
            //................
            IWebDriver gc = new ChromeDriver();
            
            //...loging in 
            gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/Login");
            gc.FindElement(By.Name("LoginID")).SendKeys("lightbot") ; 
            gc.FindElement(By.Name("Password")).SendKeys("RPAuser!");
            gc.FindElement(By.Name("Password")).SendKeys(Keys.Enter);
            //.............
            gc.Navigate().GoToUrl("https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\\AEE1+Ancillary+Risk+Fees");
            gc.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            gc.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
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
                    using (var db = new Entities1())
                    {
                        db.Tables.Add(table);
                        db.SaveChanges();
                    }
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
                    using (var db = new Entities1())
                    {
                        db.Tables.Add(table);
                        db.SaveChanges();
                    }
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
                    using (var db = new Entities1())
                    {
                        db.Tables.Add(table);
                        db.SaveChanges();
                    }
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
                    using (var db = new Entities1())
                    {
                        db.Tables.Add(table);
                        db.SaveChanges();
                    }
                }

            }
            //.........
            var db1 = new Entities1();

            Console.WriteLine("Data extraction done ;");
            Console.WriteLine("Data with location :");
            foreach (Table i in db1.Tables.Where(s => s.location == null))
            {
                //gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next");
                //System.Threading.Thread.Sleep(5000);
                //gc.FindElements(By.ClassName("cs-header-module-item"))[2].Click();
                Program test = new Program();
                //test.Create(gc,i);
            }
            var delimiter = "\t";
            using (var writer = new System.IO.StreamWriter("hello.txt"))
            {
                foreach (Table i in db1.Tables.Where(s => s.location != null))
                {
                    writer.WriteLine(i.Case_no + delimiter + i.billDate + delimiter + i.eventCode + delimiter + i.billRates + delimiter + i.billUnits + delimiter + i.clientID + delimiter + i.location);
                }
            }

        }
        
    }
}
