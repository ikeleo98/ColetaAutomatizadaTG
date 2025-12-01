using ColetaAutomatizadaTG.Interfaces;
using ColetaAutomatizadaTG.Services;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace ColetaAutomatizadaTG;

public class Worker : BackgroundService
{
    private readonly ILogger<Worker> _logger;
    private readonly ISeleniumHandler _seleniumHandler;
    private readonly ExcelExporter _excelExporter;
    private readonly IAutomacaoRepository _automacaoRepository;

    public Worker(
        ILogger<Worker> logger,
        ISeleniumHandler seleniumHandler,
        ExcelExporter excelExporter,
        IAutomacaoRepository automacaoRepository)
    {
        _logger = logger;
        _seleniumHandler = seleniumHandler;
        _excelExporter = excelExporter;
        _automacaoRepository = automacaoRepository;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            _seleniumHandler.CriarDriver();

            try
            {
                _logger.LogInformation("Iniciando fluxo de coleta...");

                var parametros = await _automacaoRepository.ObterParametroAtivosAsync(stoppingToken);
                if (parametros == null)
                {
                    _logger.LogWarning("Nenhum parâmetro de coleta ativo encontrado.");
                    return;
                }

                var dataExecucao = DateTime.UtcNow;

                foreach (var parametro in parametros)
                {
                    try
                    {
                        var moeda = (parametro.TipoMoeda ?? string.Empty).Trim();

                        if (string.IsNullOrWhiteSpace(moeda))
                            moeda = "SEM_MOEDA";

                        var moedaSanitizada = moeda.Replace(" ", "_").Replace("/", "-").ToUpperInvariant();
                        var pasta = @"C:\ArquivosCotacoes";

                        Directory.CreateDirectory(pasta);

                        var caminho = Path.Combine(
                            pasta,
                            $"Cotacoes_{dataExecucao:yyyyMMdd}_{moedaSanitizada}.xlsx");

                        var cotacoes = _seleniumHandler.ObterCotacoesAsync(parametro.DataInicial, parametro.DataFinal, parametro.TipoMoeda);

                        _excelExporter.ExportarCotacoesParaExcel(cotacoes, caminho);

                        var loteId = await _automacaoRepository.RegistrarLoteAsync(parametro, dataExecucao, cotacoes.Count, true, null, stoppingToken);

                        await _automacaoRepository.InserirCotacoesAsync(loteId, cotacoes, parametro.TipoMoeda, stoppingToken);

                        _logger.LogInformation(
                            "Fluxo concluído para moeda {Moeda}. Excel em {Caminho}",
                            parametro.TipoMoeda, caminho);
                    }
                    catch (IOException ioEx) when (ioEx.Message.Contains("mesmo tamanho"))
                    {
                        var pasta = @"C:\ArquivosCotacoes";
                        var caminhoArquivo = Path.Combine(
                            pasta,
                            $"Cotacoes_{DateTime.UtcNow:yyyyMMdd}_{parametro.TipoMoeda}.xlsx");

                        _logger.LogWarning(
                            "Arquivo já existe e possui o mesmo tamanho. Ignorando parâmetro da moeda {Moeda}. Caminho: {Caminho}",
                            parametro.TipoMoeda,
                            caminhoArquivo
                        );

                        continue;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(
                            ex,
                            "Erro ao processar parâmetro da moeda {Moeda}. Parâmetro será ignorado.",
                            parametro.TipoMoeda);

                        continue;
                    }
                }
            }
            catch (Exception e)
            {

                throw;
            }
            finally
            {
                _seleniumHandler.FecharDriver();
            }

            await Task.Delay(TimeSpan.FromMinutes(1), stoppingToken);
        }
    }

}
