# OBIM Orça

Site para montar orçamentos de obras residenciais, passo a passo, com explicação em cada campo e a base SINAPI atualizada sozinha todo mês.

## O que tem aqui

| Arquivo | Para que serve |
| --- | --- |
| `index.html` | Assistente guiado: obra, local e base, desoneração, encargos, BDI, etapas e resumo |
| `orcamento.html` | Lançamento dos serviços: importa os PDFs do CYPECAD, sugere a composição SINAPI de cada item, avisa o que costuma faltar e exporta para Excel |
| `busca.html` | Consulta às composições e insumos da SINAPI por estado |
| `dados/sinapi/` | Base SINAPI convertida para o site (gerada pelo robô, não edite à mão) |
| `scripts/atualizar_sinapi.py` | Robô que baixa a SINAPI da Caixa e converte |
| `.github/workflows/publicar.yml` | Roda o robô nos dias 12, 20 e 28 de cada mês e publica o site |

## Como ativar (uma vez só)

1. No GitHub, abra **Settings → Pages** e em **Source** escolha **GitHub Actions**.
2. Abra **Actions → Atualizar SINAPI e publicar site → Run workflow** para baixar a primeira base.
3. O endereço do site aparece em **Settings → Pages** (algo como `https://ewertonhenrique.github.io/orca-facil/`).

Depois disso, o robô roda sozinho. Se a Caixa mudar o formato da planilha e o robô falhar, o GitHub manda um e-mail avisando, e o site continua com a última base que deu certo.

## Fonte dos dados

SINAPI, Caixa Econômica Federal e IBGE: <https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/>. Faixas de BDI conforme o Acórdão TCU 2622/2013.
