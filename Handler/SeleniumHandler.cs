using ColetaAutomatizadaTG.Interfaces;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace ColetaAutomatizadaTG.Handler;

public class SeleniumHandler : ISeleniumHandler
{
    private IWebDriver? _driver;

    public SeleniumHandler()
    {
    }

    public void CriarDriver()
    {
        if (_driver != null)
            return;

        var options = new ChromeOptions();

        //options.AddArgument("--headless=new");
        options.AddArgument("--window-size=1920,1080");
        options.AddArgument("--disable-gpu");
        options.AddArgument("--no-sandbox");
        options.AddArgument("--disable-dev-shm-usage");

        _driver = new ChromeDriver(options);
    }

    public List<CotacaoDto> ObterCotacoesAsync(DateTime dataInicial, DateTime dataFinal, string moeda)
    {
        try
        {
            CriarDriver(); // garante que o driver existe

            _driver.Navigate().GoToUrl("https://www.bcb.gov.br/estabilidadefinanceira/historicocotacoes");

            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(15));

            wait.Until(drv =>
            {
                try
                {
                    var iframe = drv.FindElements(By.TagName("iframe"))
                                    .FirstOrDefault(f =>
                                        (f.GetAttribute("src") ?? "")
                                            .Contains("consultaBoletim"));

                    if (iframe == null)
                        return false;

                    drv.SwitchTo().Frame(iframe);
                    return true;
                }
                catch
                {
                    return false;
                }
            });

            string dataIniTexto = dataInicial.ToString("ddMMyyyy");
            string dataFimTexto = dataFinal.ToString("ddMMyyyy");

            var campoDataIni = _driver.FindElement(By.Id("DATAINI")); 
            campoDataIni.Clear();
            campoDataIni.SendKeys(dataIniTexto);

            var campoDataFim = _driver.FindElement(By.Id("DATAFIM")); 
            campoDataFim.Clear();
            campoDataFim.SendKeys(dataFimTexto);

            var selectMoedaElement = _driver.FindElement(By.Name("ChkMoeda"));
            var selectMoeda = new SelectElement(selectMoedaElement);

            selectMoeda.SelectByText(moeda);


            var botaoPesquisar = _driver.FindElement(By.CssSelector("input[type='submit'][value='Pesquisar']"));
            botaoPesquisar.Click();

            Thread.Sleep(1000);

            var cotacoes = ObterCotacoesTabela();

            _driver.SwitchTo().DefaultContent();

            return cotacoes;
        }
        catch (Exception e)
        {

            throw;
        }
    }

    public List<CotacaoDto> ObterCotacoesTabela()
    {
        var resultado = new List<CotacaoDto>();
        var tabela = _driver.FindElement(By.XPath("(//table)[1]"));

        var linhas = tabela.FindElements(By.XPath(".//tr[td]"));

        foreach (var linha in linhas)
        {
            var colunas = linha.FindElements(By.TagName("td"));
            if (colunas.Count < 4)
                continue;

            var dto = new CotacaoDto
            {
                Data = colunas[0].Text.Trim(),
                ValorCompra = colunas[2].Text.Trim(),
                ValorVenda = colunas[3].Text.Trim()
            };

            resultado.Add(dto);
        }

        return resultado;
    }

    public void FecharDriver()
    {
        try
        {
            if (_driver != null)
            {
                _driver.Quit();
                _driver.Dispose();
                _driver = null;
            }
        }
        catch
        {

        }
    }

}
