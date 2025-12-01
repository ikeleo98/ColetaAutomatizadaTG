# ColetaAutomatizadaTG

Sistema desenvolvido para automação da coleta de cotações do Banco Central do Brasil, utilizando RPA com Selenium WebDriver, persistência em SQL Server, geração de relatórios em Excel e execução contínua como serviço do Windows.

## Descrição do Projeto
O ColetaAutomatizadaTG automatiza o acesso ao portal Histórico de Cotações do Banco Central, executa consultas parametrizadas, extrai os dados, valida, padroniza e armazena tudo em banco de dados relacional. A aplicação roda como serviço do Windows.

## Tecnologias Utilizadas
- .NET 8 (Worker Service)
- Selenium WebDriver + ChromeDriver
- SQL Server
- Dapper
- EPPlus
- Serilog
- Windows Service
- PowerShell

## Publicação
```
dotnet publish -c Release
```
Copiar os binários para:
```
C:\RPA\ColetaAutomatizadaTG
```

## Instalação do Serviço
```
New-Service -Name "RPA.ColetaAutomatizadaTG" -BinaryPathName "C:\RPA\ColetaAutomatizadaTG\ColetaAutomatizadaTG.exe" -StartupType Automatic
Start-Service -Name "RPA.ColetaAutomatizadaTG"
```

## Validações Implementadas
- Parâmetros ativos
- Prevenção de duplicidade
- Normalização
- Validação do Excel
- Logs estruturados

## Saídas
Excel em:
```
C:\ArquivosCotacoes
```

## Autor
Leonardo Toshio Ikehara

## Licença
Projeto acadêmico (TG).
