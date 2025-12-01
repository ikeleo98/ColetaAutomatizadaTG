using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace ColetaAutomatizadaTG.Interfaces
{
    public interface ISeleniumHandler
    {
        List<CotacaoDto> ObterCotacoesAsync(DateTime dataInicial, DateTime dataFinal, string moeda);

        void CriarDriver();

        void FecharDriver();
    }
}
