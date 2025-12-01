namespace ColetaAutomatizadaTG.Interfaces;

public interface IAutomacaoRepository
{
    Task<IEnumerable<ParametroColeta>> ObterParametroAtivosAsync(CancellationToken cancellationToken = default);

    Task<int> RegistrarLoteAsync(ParametroColeta parametro,DateTime dataExecucao,int quantidade,bool sucesso,string? mensagemErro,CancellationToken cancellationToken = default);

    Task InserirCotacoesAsync(int loteCotacaoId,IEnumerable<CotacaoDto> cotacoes,string tipoMoeda,CancellationToken cancellationToken = default);
}
