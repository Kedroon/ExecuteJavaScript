using System;
using System.Collections.Generic;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.PhantomJS;
using System.Speech.Synthesis;
using System.IO;
using OpenQA.Selenium.Support.UI;
using System.Data.OleDb;
using System.Threading;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using OpenQA.Selenium.Remote;
using System.Drawing.Imaging;
using HtmlAgilityPack;
using System.Net;
using System.Diagnostics;
using System.Drawing;
using System.Collections.Concurrent;

namespace PMMBotMySql
{
    class Program
    {
        static ChromeDriver js;
        static OleDbConnection connection;
        static string query;
        static List<bool> cnpjcpfValidos = new List<bool>();
        static string excelpath = @"C:\TempExcel\rel_notas_aceite_" + DateTime.Now.ToString("MM") + "-" + DateTime.Now.ToString("yyyy") + ".xls";
        static int superterminaiscount = 0;
        static int auroraeadicount = 0;
        static int chibataocount = 0;
        static int notasvalidas = 0;
        static BlockingCollection<Client> ClientQueue = new BlockingCollection<Client>(10);
        static ConcurrentDictionary<int,Uri> urls_to_download;


        static void Main(string[] args)
        {
            connection = new OleDbConnection();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\migue\OneDrive\Documentos\Notas.accdb;
Persist Security Info=False;";
            // speak();

            automation();

        }
        static private void clickVirtualButton(string num, ChromeDriver js)
        {
            js.FindElementByXPath("//img[contains(@src,'/images/teclado/tec_" + num + ".gif')]").Click();

        }



        static void automation()
        {
            /*     ReadNote note = new ReadNote("oi");
                 note.StartAnalysis();
                 Console.Read();*/

            //Environment.SetEnvironmentVariable("webdriver.phantomjs.driver", "phantomjs.exe");
            Environment.SetEnvironmentVariable("webdriver.chrome.driver", "chromedriver.exe");
            ChromeOptions options = new ChromeOptions();
            options.AddUserProfilePreference("download.default_directory", @"C:\TempExcel");
            //options.AddArgument("--no-startup-window");

            Console.WriteLine("Iniciando Chrome");

            js = new ChromeDriver(options);
            OpenQA.Selenium.Cookie cookie1 = new OpenQA.Selenium.Cookie("PID", "2524");
            OpenQA.Selenium.Cookie cookie2 = new OpenQA.Selenium.Cookie("MOBI", "560801");
            OpenQA.Selenium.Cookie cookie3 = new OpenQA.Selenium.Cookie("TIPO", "0");
            OpenQA.Selenium.Cookie cookie4 = new OpenQA.Selenium.Cookie("SUBTIPO", "");

            //Console.WriteLine("Profile do firefox nao encontrado");





            Console.WriteLine("Pagina 1");
            js.Navigate().GoToUrl("https://www3.gissonline.com.br/interna/default.cfms");
            js.Manage().Cookies.AddCookie(cookie1);
            js.Manage().Cookies.AddCookie(cookie2);
            js.Manage().Cookies.AddCookie(cookie3);
            js.Manage().Cookies.AddCookie(cookie4);
            bool page2 = false;
        Page1:
            try
            {
                js.Navigate().GoToUrl("https://www3.gissonline.com.br/interna/default.cfm");
                js.SwitchTo().Frame(0);
                js.FindElementByXPath("//img[contains(@src,'images/bt_menu__06_off.jpg')]").Click();
            }

            catch (UnhandledAlertException)
            {
                js.SwitchTo().Alert().Accept();
                js.SwitchTo().DefaultContent();
                goto Page1;
            }
            catch (Exception err)
            {


                Console.WriteLine(err.Message);
                //js.Navigate().Refresh();
                goto Page1;

            }
            
        Page2:
            Console.WriteLine("Pagina 2");
            try
            {
                js.SwitchTo().DefaultContent();
                js.SwitchTo().Frame(2);
                DateTime time = DateTime.Now;
                if (!page2)
                {
                    js.FindElement(By.Name("mes")).SendKeys(time.ToString("MM"));
                    js.FindElement(By.Name("ano")).SendKeys(time.Year.ToString());
                    page2 = true;
                }
                else
                {
                    js.FindElement(By.Name("ano")).SendKeys(" ");
                    js.FindElement(By.Name("ano")).Click();
                    Console.WriteLine("oi");
                }
                js.FindElement(By.LinkText("Notas Recebidas")).Click();
            }
            catch (UnhandledAlertException)
            {
                js.SwitchTo().Alert().Accept();
                js.SwitchTo().DefaultContent();
                goto Page1;
            }

            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                //js.Navigate().Refresh();
                goto Page1;
            }
            int count = 0;
            ReadOnlyCollection<IWebElement> element;
            string mwh;
            bool first = true;
        Page3:
            Console.WriteLine("Pagina 3");
            try
            {
                js.SwitchTo().DefaultContent();
                js.SwitchTo().Frame(2);
                new SelectElement(js.FindElement(By.Name("maxrow"))).SelectByText("500");

                if (File.Exists(excelpath))
                {
                    File.Delete(excelpath);
                }

                js.FindElementByXPath("//a[contains(text(),'GERAR ARQUIVO EXCEL')]").Click();
                mwh = js.CurrentWindowHandle;

            }
            catch (UnhandledAlertException)
            {
                js.SwitchTo().Alert().Accept();
                js.SwitchTo().DefaultContent();
                goto Page1;
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                //js.Navigate().Refresh();
                goto Page1;
            }
            //Esperar o termino do download da planilha
            Console.WriteLine("Downloading Planilha");
            for (var i = 0; i < 30; i++)
            {
                if (File.Exists(excelpath))
                {
                    break;
                }
                Thread.Sleep(1000);
            }
            long length;
        FileLength:
            try
            {
                length = new FileInfo(excelpath).Length;
            }
            catch (Exception)
            {

                goto FileLength;
            }

            for (var i = 0; i < 30; i++)
            {
                Thread.Sleep(1000);
                var newLength = new FileInfo(excelpath).Length;
                if (newLength == length && length != 0) { break; }
                length = newLength;
            }
            Console.WriteLine("Download concluido");
            Console.WriteLine("Analisando planilha");
            ListOfCNPJCPF(); //Analisar planilha
            Thread.Sleep(3000);
            Console.WriteLine("Analise concluida");
            //speak();
            Console.WriteLine(cnpjcpfValidos.Count);
            //MessageBox.Show(cnpjcpfValidos.Count.ToString());

            /*  var cookies = js.Manage().Cookies;
              ReadOnlyCollection<OpenQA.Selenium.Cookie> wtf = cookies.AllCookies;
              foreach (var item in wtf)
              {   
                  Console.WriteLine(item.Name +" "+ item.Value);
              }*/
            
            HtmlAgilityPack.HtmlDocument page = new HtmlAgilityPack.HtmlDocument();
            page.LoadHtml(js.PageSource);
            //Console.WriteLine(js.PageSource);
            List<HtmlNode> urlOfNotas = new List<HtmlNode>();
            foreach (var item in page.DocumentNode.SelectNodes("//a[starts-with(@onclick,'janela')]"))
            {
                urlOfNotas.Add(item);
            }
            Console.WriteLine(urlOfNotas.Count);

            int loopPaginas=500;
            while (cnpjcpfValidos.Count>loopPaginas)
            {
                int tempvalue = loopPaginas+1;
                js.FindElementByXPath("//a[contains(@onclick,'document.formPag.startrow.value="+tempvalue+";document.formPag.submit();')]").Click();
                HtmlAgilityPack.HtmlDocument nextpage = new HtmlAgilityPack.HtmlDocument();
                nextpage.LoadHtml(js.PageSource);
                foreach (var item in nextpage.DocumentNode.SelectNodes("//a[starts-with(@onclick,'janela')]"))
                {
                    urlOfNotas.Add(item);
                    
                }
                
                loopPaginas += 500;

            }

            
            string temp;

            urls_to_download = new ConcurrentDictionary<int, Uri>();
            foreach (var item in urlOfNotas)
            {
                if (cnpjcpfValidos[count] == true)
                {
                    temp = item.OuterHtml;
                    //Console.WriteLine(@"',430,260)""><img src=""../biblioteca/images/PL_FindResults_R.png"" title=""Dados da nota fiscal"" border=""0""></a>");
                    temp = temp.Replace(@"<a href=""javascript:;"" onclick=""janela('..", "").Replace(@"',430,260)""><img src=""../biblioteca/images/PL_FindResults_R.png"" title=""Dados da nota fiscal"" border=""0""></a>", "");
                    temp = "https://www3.gissonline.com.br" + temp;
                    temp = temp.Replace("amp;", "");
                    Console.WriteLine(temp);
                    urls_to_download[urls_to_download.Count] = new Uri(temp);
                    //urls_to_download.Add(new Uri(temp));
                    count++;
                }

            }
            Console.WriteLine(count);
            List<Client> clients = new List<Client>();
            for (int i = 0; i < 10; i++)
            {
                Client client = new Client();
                client.Headers.Add(HttpRequestHeader.Cookie,
              "PID=2524;" +
              "MOBI=560801"
              );
                client.DownloadStringCompleted += (sender, e) => Web_DownloadStringCompleted(sender, e, client);
                clients.Add(client);
            }

            foreach (var item in clients)
            {
                ClientQueue.Add(item);
            }
            var watch = System.Diagnostics.Stopwatch.StartNew();
            int urlatual=0;

            while (urls_to_download.Count>urlatual)
            {
                Console.WriteLine(urls_to_download.Count);
                var worker = ClientQueue.Take();
            DownloadString:
                try
                {
                    worker.url = urls_to_download[urlatual];
                    worker.DownloadStringAsync(worker.url);
                }
                catch (Exception)
                {
                    goto DownloadString;

                }
                urlatual++;
            }
                //Console.WriteLine(url);
                

                //count++;
                //Console.WriteLine(count);
            



            //loop para abrir as notas

            watch.Stop();
            Console.WriteLine("Execution Time: " + (watch.ElapsedMilliseconds / 1000) + "Seconds");
            // Console.ReadLine();
            Thread.Sleep(2000);
            MessageBox.Show("Execution Time: " + (watch.ElapsedMilliseconds / 1000) + "Seconds");

        }

        static private void ListOfCNPJCPF()
        {

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            string workbookPath = @"C:\TempExcel\rel_notas_aceite_" + DateTime.Now.ToString("MM") + "-" + DateTime.Now.ToString("yyyy");
            var workbooks = excelApp.Workbooks;
            Excel.Workbook excelWorkbook = workbooks.Open(workbookPath,
                    0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);
            Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            string currentSheet = "rel_notas_aceite_" + DateTime.Now.ToString("MM") + "-" + DateTime.Now.ToString("yyyy");
            //MessageBox.Show(currentSheet);
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
            int i = 4;
            while (excelWorksheet.Cells[i, 11].Value != null)
            {
                string cnpjcpf = excelWorksheet.Cells[i, 11].Value2.ToString();
                if (cnpjcpf == "84098383000172")
                {
                    cnpjcpfValidos.Add(true);
                    chibataocount++;
                    notasvalidas++;
                    //Console.WriteLine(excelWorksheet.Cells[i, 11].Value2);
                }
                else if (cnpjcpf == "4694548000130")
                {
                    cnpjcpfValidos.Add(true);
                    auroraeadicount++;
                    notasvalidas++;
                }
                else if (cnpjcpf == "4335535000255")
                {
                    cnpjcpfValidos.Add(true);
                    superterminaiscount++;
                    notasvalidas++;
                }
                else
                {
                    cnpjcpfValidos.Add(true);
                }

                i++;
            }
            Marshal.ReleaseComObject(excelWorksheet);
            Marshal.ReleaseComObject(excelSheets);
            Marshal.ReleaseComObject(excelWorkbook);
            Marshal.ReleaseComObject(workbooks);
            excelApp.Quit();

        }
        static void speak()
        {
            SpeechSynthesizer synthesizer;
            synthesizer = new SpeechSynthesizer();
            synthesizer.Rate = 0;
            synthesizer.SelectVoice("Microsoft Maria Desktop");
            synthesizer.SpeakAsync("Senhor Mestre do Universo, eu encontrei " + notasvalidas + " notas fiscais. Sendo " + superterminaiscount + " do Super Terminais, " + auroraeadicount + " da Aurora Eadi, e " + chibataocount + " do Chibatão.");

        }

        private static void Web_DownloadStringCompleted(object sender, DownloadStringCompletedEventArgs e, Client client)
        {
            ReadNote note;

            try
            {
                note = new ReadNote(e.Result , client.url);
            }
            catch (Exception)
            {

                ClientQueue.Add(client);
                urls_to_download[urls_to_download.Count] = client.url;
                //urls_to_download.Add(client.url);
                return;
            }

            note.StartAnalysis();
            ClientQueue.Add(client);
            


        }

    }





}