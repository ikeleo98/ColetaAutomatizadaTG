using System.Data;
using System.Globalization;
using ColetaAutomatizadaTG.Interfaces;
using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace ColetaAutomatizadaTG.Data;

public class AutomacaoRepository : IAutomacaoRepository
{
    private readonly string _connectionString;

    public AutomacaoRepository(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("SqlServer")
                           ?? throw new InvalidOperationException("ConnectionStrings:SqlServer não configurada.");
    }

    public async Task<IEnumerable<ParametroColeta>> ObterParametroAtivosAsync(
    CancellationToken cancellationToken = default)
    {
        try
        {
            const string sql = @"
            SELECT 
                ParametroColetaId,
                DataInicial,
                DataFinal,
                TipoMoeda,
                Ativo,
                DataInclusao
            FROM ParametroColeta
            WHERE Ativo = 1
            ORDER BY DataInclusao DESC;";

            await using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync(cancellationToken);

            var parametros = await conn.QueryAsync<ParametroColeta>(sql);

            return parametros;
        }
        catch (Exception e)
        {
            throw;
        }
    }


    public async Task<int> RegistrarLoteAsync(ParametroColeta parametro, DateTime dataExecucao, int quantidade, bool sucesso, string? mensagemErro, CancellationToken cancellationToken = default)
    {
        try
        {
            const string sql = @"
            INSERT INTO LoteCotacao
            (
                ParametroColetaId,
                DataExecucao,
                DataInicial,
                DataFinal,
                TipoMoeda,
                Quantidade,
                Sucesso,
                MensagemErro
            )
            OUTPUT INSERTED.LoteCotacaoId
            VALUES
            (
                @ParametroColetaId,
                @DataExecucao,
                @DataInicial,
                @DataFinal,
                @TipoMoeda,
                @Quantidade,
                @Sucesso,
                @MensagemErro
            );";

            var parametros = new
            {
                parametro.ParametroColetaId,
                DataExecucao = dataExecucao,
                parametro.DataInicial,
                parametro.DataFinal,
                parametro.TipoMoeda,
                Quantidade = quantidade,
                Sucesso = sucesso,
                MensagemErro = (object?)mensagemErro ?? DBNull.Value
            };

            await using var conn = new SqlConnection(_connectionString);
            await conn.OpenAsync(cancellationToken);

            var loteId = await conn.ExecuteScalarAsync<int>(sql, parametros);
            return loteId;
        }
        catch (Exception e)
        {

            throw;
        }
    }

    public async Task InserirCotacoesAsync(int loteCotacaoId, IEnumerable<CotacaoDto> cotacoes, string tipoMoeda, CancellationToken cancellationToken = default)
    {
        const string sql = @"
            INSERT INTO Cotacao
            (
                LoteCotacaoId,
                DataCotacao,
                TipoMoeda,
                ValorCompra,
                ValorVenda,
                DataInclusao
            )
            VALUES
            (
                @LoteCotacaoId,
                @DataCotacao,
                @TipoMoeda,
                @ValorCompra,
                @ValorVenda,
                SYSUTCDATETIME()
            );";

        await using var conn = new SqlConnection(_connectionString);
        await conn.OpenAsync(cancellationToken);

        using var trans = await conn.BeginTransactionAsync(cancellationToken);

        try
        {
            foreach (var c in cotacoes)
            {
                if (!DateTime.TryParse(c.Data, out var dataCotacao))
                    continue;

                var valorCompra = ParseDecimalPtBr(c.ValorCompra);
                var valorVenda = ParseDecimalPtBr(c.ValorVenda);

                var param = new
                {
                    LoteCotacaoId = loteCotacaoId,
                    DataCotacao = dataCotacao,
                    TipoMoeda = string.IsNullOrWhiteSpace(c.TipoMoeda) ? tipoMoeda : c.TipoMoeda,
                    ValorCompra = valorCompra,
                    ValorVenda = valorVenda
                };

                await conn.ExecuteAsync(sql, param, trans);
            }

            await trans.CommitAsync(cancellationToken);
        }
        catch
        {
            await trans.RollbackAsync(cancellationToken);
            throw;
        }
    }

    private static decimal ParseDecimalPtBr(string valorTexto)
    {
        if (string.IsNullOrWhiteSpace(valorTexto))
            return 0m;

        valorTexto = valorTexto.Trim();

        if (decimal.TryParse(valorTexto, NumberStyles.Any, new CultureInfo("pt-BR"), out var valor))
            return valor;

        valorTexto = valorTexto.Replace(',', '.');
        if (decimal.TryParse(valorTexto, NumberStyles.Any, CultureInfo.InvariantCulture, out valor))
            return valor;

        return 0m;
    }
}
