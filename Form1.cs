using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;

namespace googlesearch
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;

            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("headless");

            IWebDriver driver = new ChromeDriver(chromeDriverService,chromeOptions);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            driver.Manage().Window.Minimize();

            Application.UseWaitCursor = true;
            try
            {                
                string searchword = "台積電";
                //driver.Manage().Window.Size = new Size(1280, 720);
                //string url = Uri.EscapeUriString("https://www.google.com/search?q=台積電&tbm=nws&start=0");
                //driver.Navigate().GoToUrl(url);

                int page_number = 1;
                int start_no = 0;
                string go_next = "Y";
                textBox1.Clear();
                while (go_next == "Y")
                {
                    start_no = (page_number - 1) * 10;
                    //string url = Uri.EscapeUriString("https://www.google.com/search?q=" + searchword + "&tbm=nws&Ccd_min:4/1/2020&Ccd_max:6/14/2020&start=" + start_no);
                    string url = Uri.EscapeUriString("https://www.google.com/search?q=" + searchword + "&tbm=nws&tbs=qdr:y&start=" + start_no);
                    driver.Navigate().GoToUrl(url);

                    //讀取"result status"
                    if (driver.PageSource.IndexOf("找不到符合") != -1)
                    {
                        go_next = "N";
                    }
                    else
                    {
                        textBox1.AppendText("===第" + page_number.ToString() + "頁===" + Environment.NewLine);
                        //寫入新聞資訊                    
                        string xpath_str = "";
                        IList<IWebElement> g = driver.FindElements(By.ClassName("g"));
                        textBox1.AppendText("===共" + g.Count.ToString() + "條===" + Environment.NewLine);
                        for (int i = 1; i <= g.Count; i++)
                        {
                            textBox1.AppendText("第" + i.ToString() + "條" + Environment.NewLine);
                            //title
                            xpath_str = "//*[@id='rso']/div[" + i + "]/div/div[1]/h3/a";
                            IWebElement a = driver.FindElement(By.XPath(xpath_str));
                            textBox1.AppendText("Title:" + a.Text + Environment.NewLine);
                            //link
                            textBox1.AppendText("link:" + a.GetAttribute("href") + Environment.NewLine);
                            //media
                            xpath_str = "//*[@id='rso']/div[" + i + "]/div/div[1]/div[1]/span[1]";
                            IWebElement a1 = driver.FindElement(By.XPath(xpath_str));
                            textBox1.AppendText("media:" + a1.Text + Environment.NewLine);
                            //pubdate
                            xpath_str = "//*[@id='rso']/div[" + i + "]/div/div[1]/div[1]/span[3]";
                            IWebElement a2 = driver.FindElement(By.XPath(xpath_str));
                            // 進行當日發布新聞日期調整
                            if (a2.Text.IndexOf("前") > -1)
                            {
                                textBox1.AppendText("date:" + DateTime.Today.ToString("yyyy年MM月dd日") + Environment.NewLine);
                            }
                            else
                            {
                                textBox1.AppendText("date:" + a2.Text + Environment.NewLine);
                            }                                
                            //desc
                            xpath_str = "//*[@id='rso']/div[" + i + "]/div/div/div[2]";
                            IWebElement a3 = driver.FindElement(By.XPath(xpath_str));
                            textBox1.AppendText("desc:" + a3.Text + Environment.NewLine + Environment.NewLine);

                            //附掛新聞: YiHbdc
                            xpath_str = "//*[@id='rso']/div["+i+"]";
                            IList<IWebElement> g1 = driver.FindElement(By.XPath(xpath_str)).FindElements(By.ClassName("YiHbdc"));
                            for (int j = 1 ; j <= g1.Count; j++)
                            {
                                textBox1.AppendText("第" + i.ToString() + "條之 "+j.ToString() + Environment.NewLine);
                                //title
                                xpath_str = "//*[@id='rso']/div[" + i + "]/div/div["+(j*2)+"]/a";
                                IWebElement card = driver.FindElement(By.XPath(xpath_str));                              
                                textBox1.AppendText("Title:" + card.Text + Environment.NewLine);
                                //link
                                textBox1.AppendText("link:" + card.GetAttribute("href") + Environment.NewLine);
                                //media
                                xpath_str = "//*[@id='rso']/div[" + i + "]/div/div[" + (j * 2) + "]/span[1]";
                                IWebElement card1 = driver.FindElement(By.XPath(xpath_str));
                                textBox1.AppendText("media:" + card1.Text + Environment.NewLine);
                                //pubdate
                                xpath_str = "//*[@id='rso']/div[" + i + "]/div/div[" + (j * 2) + "]/span[3]";
                                IWebElement card2 = driver.FindElement(By.XPath(xpath_str));
                                // 進行當日發布新聞日期調整
                                if (card2.Text.IndexOf("前") > -1)
                                {
                                    textBox1.AppendText("date:" + DateTime.Today.ToString("yyyy年MM月dd日") + Environment.NewLine);
                                }
                                else
                                {
                                    textBox1.AppendText("pubdate:" + card2.Text + Environment.NewLine);
                                }                                
                                //desc
                                textBox1.AppendText("desc:" + card.Text + Environment.NewLine + Environment.NewLine);
                            }

                            //附掛新聞: ErI7Gd
                            xpath_str = "//*[@id='rso']/div[" + i + "]";
                            IList<IWebElement> g2 = driver.FindElement(By.XPath(xpath_str)).FindElements(By.ClassName("ErI7Gd"));
                            for (int k = 1; k <= g2.Count; k++)
                            {
                                textBox1.AppendText("第" + i.ToString() + "條之 " + (k+ g1.Count).ToString() + Environment.NewLine);
                                //title
                                xpath_str = "//*[@id='rso']/div[" + i + "]/div/div[" + ((k + g1.Count) * 2) + "]/a";
                                IWebElement card = driver.FindElement(By.XPath(xpath_str));
                                textBox1.AppendText("Title:" + card.Text + Environment.NewLine);
                                //link
                                textBox1.AppendText("link:" + card.GetAttribute("href") + Environment.NewLine);
                                //media
                                xpath_str = "//*[@id='rso']/div[" + i + "]/div/div[" + ((k + g1.Count) * 2) + "]/span[1]";
                                IWebElement card1 = driver.FindElement(By.XPath(xpath_str));
                                textBox1.AppendText("media:" + card1.Text + Environment.NewLine);
                                //pubdate
                                xpath_str = "//*[@id='rso']/div[" + i + "]/div/div[" + ((k + g1.Count) * 2) + "]/span[3]";
                                IWebElement card2 = driver.FindElement(By.XPath(xpath_str));
                                // 進行當日發布新聞日期調整
                                if (card2.Text.IndexOf("前") > -1)
                                {
                                    textBox1.AppendText("date:" + DateTime.Today.ToString("yyyy年MM月dd日") + Environment.NewLine);
                                }
                                else
                                {
                                    textBox1.AppendText("pubdate:" + card2.Text + Environment.NewLine);
                                }
                                //desc
                                textBox1.AppendText("desc:" + card.Text + Environment.NewLine + Environment.NewLine);
                            }

                            Application.DoEvents();
                        }
                        page_number++;
                        go_next = "N";
                    }                    
                } ;
                textBox1.AppendText("[" + searchword + "]查詢結束!!" + Environment.NewLine);
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);                
            }
            finally
            {
                driver.Quit();
                driver.Dispose();
                Application.UseWaitCursor = false;
            }
                
            //}  );
        }
    }
}
