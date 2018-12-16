using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support;
using ClosedXML.Excel;
using System.Collections;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;

namespace PageObjectLinkedin.Testes
{
    [TestClass]
    public class PesquisaLinkedinTeste
    {
        IWebDriver driver = new ChromeDriver();
        PesquisaLinkedin linkedin = new PesquisaLinkedin();
        string UrlSite = "https://www.linkedin.com/";

        [TestInitialize]
        public void setUp()
        {
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(UrlSite);
            driver.Manage().Cookies.DeleteAllCookies();

        }

        [TestMethod]
        public void setPesquisaLinkedin()
        {
            var wb = new XLWorkbook(@"D:\Planilha\dados.xlsx");
            var planilha = wb.Worksheet(1);
            Actions action = new Actions(driver);


            driver.FindElement(linkedin.loginComboBox).SendKeys(planilha.Cell("A" + 2).Value.ToString());
            driver.FindElement(linkedin.passWordComboBox).SendKeys(planilha.Cell("B" + 2).Value.ToString());
            driver.FindElement(linkedin.clickSignIN).Click();

            if (driver.Url == "https://www.linkedin.com/feed/")
            {

                System.Threading.Thread.Sleep(5000);
                driver.FindElement(linkedin.pesquisaPessoa).SendKeys(planilha.Cell("C" + 2).Value.ToString());
                driver.FindElement(linkedin.pesquisaPessoa).SendKeys(Keys.Enter);

                System.Threading.Thread.Sleep(3000);

                driver.FindElement(linkedin.clickConexoes).Click();

                WebDriverWait espera = new WebDriverWait(driver, new TimeSpan(0, 0, 4));
                IWebElement clickCheckBox = espera.Until(ExpectedConditions.ElementToBeClickable(driver.FindElement(linkedin.clickCheckBox)));
                clickCheckBox.Click();

                driver.FindElement(linkedin.clickAplicar).Click();

                System.Threading.Thread.Sleep(5000);

                IList<IWebElement> elementos = driver.FindElements(linkedin.listaElementosConectar);

                Console.Write("Elementos encontrados " + elementos.Count);

                var allOptions = elementos.Count;

                for (int i = 0; i < allOptions; i++)

                {


                    if (elementos[i].Displayed)

                    {

                        // Clica no botão conectar
                        driver.FindElement(linkedin.clickConectar).Click();

                        // Mapeia o modal após o clique no botão conectar
                        driver.FindElement(linkedin.mapearModal);

                        // Clica no botão enviar agora
                        driver.FindElement(linkedin.clickEnviarAgora).Click();

                        // Realiza o refresh da pagina
                        driver.Navigate().Refresh();

                        // Espera pagina carregar
                        System.Threading.Thread.Sleep(4000);


                    }

                    action.SendKeys(Keys.End);
                    action.SendKeys(Keys.Home);
                }


                // Realizar o print da tela
                ITakesScreenshot screenshotDriver = driver as ITakesScreenshot;
                Screenshot screenshot = screenshotDriver.GetScreenshot();
                String fp = "D:\\Screen\\" + "_" + DateTime.Now.ToString("dd_MMMM_hh_mm_ss_tt") + ".png";
                screenshot.SaveAsFile(fp, OpenQA.Selenium.ScreenshotImageFormat.Png);


            }
            else if (driver.Url == "https://www.linkedin.com/checkpoint/lg/login?errorKey=unexpected_error")

            {


                driver.FindElement(linkedin.loginComboBoxSengundaOpcao).SendKeys(planilha.Cell("A" + 2).Value.ToString());
                driver.FindElement(linkedin.passWordComboBoxSegundaOpcao).SendKeys(planilha.Cell("B" + 2).Value.ToString());
                driver.FindElement(linkedin.clickSignINsegundaOpcao).Click();

                System.Threading.Thread.Sleep(5000);

                driver.FindElement(linkedin.pesquisaPessoa).SendKeys(planilha.Cell("C" + 2).Value.ToString());
                driver.FindElement(linkedin.pesquisaPessoa).SendKeys(Keys.Enter);

                System.Threading.Thread.Sleep(5000);

                driver.FindElement(linkedin.clickConexoes).Click();

                WebDriverWait espera = new WebDriverWait(driver, new TimeSpan(0, 0, 4));
                IWebElement clickCheckBox = espera.Until(ExpectedConditions.ElementToBeClickable(driver.FindElement(linkedin.clickCheckBox)));
                clickCheckBox.Click();

                driver.FindElement(linkedin.clickAplicar).Click();

                System.Threading.Thread.Sleep(5000);

                IList<IWebElement> elementos = driver.FindElements(linkedin.listaElementosConectar);

                Console.Write("Elementos encontrados " + elementos.Count);

                var allOptions = elementos.Count;

                for (int i = 0; i < allOptions; i++)

                {


                    if (elementos[i].Displayed)

                    {

                        // Clica no botão conectar
                        driver.FindElement(linkedin.clickConectar).Click();

                        // Mapeia o modal após o clique no botão conectar
                        driver.FindElement(linkedin.mapearModal);

                        // Clica no botão enviar agora
                        driver.FindElement(linkedin.clickEnviarAgora).Click();

                        // Realiza o refresh da pagina
                        driver.Navigate().Refresh();

                        // Espera pagina carregar
                        System.Threading.Thread.Sleep(4000);


                    }

                    action.SendKeys(Keys.End);
                    action.SendKeys(Keys.Home);
                }


                // Realizar o print da tela
                ITakesScreenshot screenshotDriver = driver as ITakesScreenshot;
                Screenshot screenshot = screenshotDriver.GetScreenshot();
                String fp = "D:\\Screen\\" + "_" + DateTime.Now.ToString("dd_MMMM_hh_mm_ss_tt") + ".png";
                screenshot.SaveAsFile(fp, OpenQA.Selenium.ScreenshotImageFormat.Png);


            }
            else

                driver.Quit();

        }



        [TestCleanup]
        public void finishTest()
        {


            driver.Quit();


        }
    }
}
