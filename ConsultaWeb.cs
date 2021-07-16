using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;

namespace CrawlerFipe
{
    public class ConsultaWeb : IDisposable
    {
        RemoteWebDriver driver;

        public ConsultaWeb()
        {
            driver = GetDriver("chrome");
        }

        public void ConsultaVeiculos()
        {
            var listaDeTipos = new string[] { "LEVE", "PESADO", "MOTOCICLO" };
            List<Veiculo> listaDeVeiculos = new List<Veiculo>();

            driver.Navigate().GoToUrl($"http://veiculos.fipe.org.br/");

            IWait<IWebDriver> wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));

            if (driver.GetElement(By.ClassName("tab-veiculos"), out IWebElement elTabVeiculos))
            {
                for (int cont = 0; cont < 3; cont++)
                {
                    var tipo = listaDeTipos[cont];
                    string divId;
                    string divId2;

                    if (cont == 0)
                    {
                        divId = "selectMarcacarro";
                        divId2 = "selectAnoModelocarro";
                    }
                    else if (cont == 1)
                    {
                        divId = "selectMarcacaminhao";
                        divId2 = "selectAnoModelocaminhao";
                    }
                    else
                    {
                        divId = "selectMarcamoto";
                        divId2 = "selectAnoModelomoto";
                    }

                    var ulTabVeiculos = elTabVeiculos.FindElements(By.TagName("ul")).First();
                    var listaDeLi = ulTabVeiculos.FindElements(By.ClassName("ilustra")).ToList();

                    var liCarro = listaDeLi[cont];
                    liCarro.Click();

                    if (driver.GetElement(By.Id(divId), out IWebElement elDropDownMarca))
                    {
                        var listaDeMarcaOptions = elDropDownMarca.FindElements(By.TagName("option")).ToList();

                        foreach (var optionMarca in listaDeMarcaOptions)
                        {
                            if (!string.IsNullOrWhiteSpace(optionMarca.Text))
                            {
                                var selectElementMarca = new SelectElement(elDropDownMarca);
                                selectElementMarca.SelectByText(optionMarca.Text);

                                if (driver.GetElement(By.Id(divId2), out IWebElement elDropDownModelo))
                                {
                                    var listaDeModeloOptions = elDropDownModelo.FindElements(By.TagName("option")).ToList();

                                    foreach (var optionModelo in listaDeModeloOptions)
                                    {
                                        if (!string.IsNullOrWhiteSpace(optionModelo.Text))
                                        {
                                            listaDeVeiculos.Add(new Veiculo { Marca = optionMarca.Text.ToUpper(), Modelo = optionModelo.Text.ToUpper(), Tipo = tipo });
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (listaDeVeiculos.Any())
            {
                var excelBytes = listaDeVeiculos.ToExcelBytes();
                File.WriteAllBytes("C:\\TabelaFipe.xlsx", excelBytes);
            }
        }

        private RemoteWebDriver GetDriver(string driver)
        {
            if (driver == "chrome")
            {
                return new ChromeDriver(AppDomain.CurrentDomain.BaseDirectory);
            }
            else if (driver == "phantom")
            {
                return new PhantomJSDriver(AppDomain.CurrentDomain.BaseDirectory);
            }
            else
            {
                throw new Exception(driver + " não implementado.");
            }
        }

        public void Dispose()
        {
            driver.Dispose();
        }
    }
}
