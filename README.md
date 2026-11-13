# Simulador de Investimentos em Fundos Imobiliários (Excel)

Desafio de Projeto da DIO: construir, a partir de uma planilha em branco, uma ferramenta em Excel que simula investimentos em fundos imobiliários — do aporte mensal ao patrimônio acumulado e aos dividendos.

## Sobre o desafio

Não há base de dados para baixar: o ponto de partida é uma planilha vazia e a planilha resolvida do Expert serve como gabarito. O objetivo é chegar, em **cinco passos**, a um simulador funcional que use as principais ferramentas do Excel aplicadas a investimentos.

A planilha final permite:

- Calcular o **patrimônio acumulado** com a função `VF` (`FV`);
- Projetar **cenários de 2 a 30 anos**;
- Usar **intervalos nomeados** para deixar as fórmulas legíveis;
- Dividir o aporte mensal por **perfil de investidor** usando `PROCV` (`VLOOKUP`) e uma **tabela de apoio**.

## Estrutura do repositório

```
.
├── README.md
├── planilha/
│   └── dio_invest.xlsx      # planilha resolvida (aba APP + tabela de apoio)
└── prints/                  # capturas de tela da ferramenta (adicionar)
```

> A pasta `prints/` deve conter as imagens das telas principais: configurações, investimento mensal, cenários e divisão por perfil.

## Como usar a planilha

1. Abra `planilha/dio_invest.xlsx`.
2. Na aba **APP**, preencha as **Configurações**:
   - **Salário**: renda mensal considerada.
   - **Rendimento da carteira**: taxa de dividendos mensal sobre o patrimônio (ex.: `0,6%`).
   - **Sugestão de investimento (30%)**: sugestão automática de aporte com base no salário.
3. Em **Investimento Mensal**, informe:
   - **Quanto investir por mês** (aporte).
   - **Por quantos anos** (horizonte do investimento).
   - **Taxa de rendimento mensal** (ex.: `1,079%`).
4. Leia o **patrimônio acumulado** e os **dividendos mensais** calculados automaticamente.
5. Consulte a tabela de **Cenários** (2, 5, 10, 20 e 30 anos) e a divisão por **Perfil**.
6. Para simular outro perfil, altere a célula **PERFIL** (`Conservador`, `Moderado` ou `Agressivo`).

## Os cinco passos da construção

1. **Organizar a entrada de dados**: montar a área de configurações (salário, rendimento da carteira, sugestão de aporte de 30%).
2. **Criar os intervalos nomeados**: nomear as células-chave da planilha
3. **Calcular patrimônio e dividendos**: usar `VF` para o patrimônio e multiplicar pelo rendimento da carteira para os dividendos mensais.
4. **Projetar os cenários**: repetir o cálculo de `VF` para horizontes de 2, 5, 10, 20 e 30 anos.
5. **Dividir o aporte por perfil**: montar a tabela de apoio com a chave `PERFIL-TIPO` e usar `PROCV` para trazer o percentual de cada tipo de FII.

## Intervalos nomeados

| Nome                  | Célula    | Significado                                  |
| --------------------- | --------- | -------------------------------------------- |
| `salario`             | `APP!D12` | Salário informado                            |
| `rendimento_carteira` | `APP!D13` | Rendimento mensal da carteira (dividendos)   |
| `sugestao_investimento` | `APP!D14` | Sugestão de aporte (30% do salário)        |
| `aporte`              | `APP!D17` | Valor investido por mês                      |
| `qtd_anos`            | `APP!D18` | Horizonte do investimento em anos            |
| `taxa_mensal`         | `APP!D19` | Taxa de rendimento mensal                    |
| `patrimonio`          | `APP!D20` | Patrimônio acumulado (resultado do `VF`)     |

## Explicação dos cálculos

### Patrimônio acumulado

```excel
=FV(taxa_mensal; qtd_anos*12; aporte*-1)
```

- `taxa_mensal`: taxa de juros por período (mensal).
- `qtd_anos*12`: número de períodos em meses.
- `aporte*-1`: o aporte é uma saída de caixa, por isso o sinal negativo.
- O tipo de pagamento é omitido (`0`), considerando aportes no **fim** de cada mês.

### Dividendos mensais

```excel
=patrimonio*rendimento_carteira
```

Os dividendos são estimados como um percentual fixo do patrimônio acumulado (ex.: `0,6%` ao mês).

### Cenários de 2 a 30 anos

| Anos | Patrimônio                            | Dividendo mensal                  |
| ---- | ------------------------------------- | --------------------------------- |
| 2    | `=FV($D$19; $A24*12; $D$17*-1)`       | `=C24*rendimento_carteira`        |
| 5    | `=FV($D$19; $A25*12; $D$17*-1)`       | `=C25*rendimento_carteira`        |
| 10   | `=FV($D$19; $A26*12; $D$17*-1)`       | `=C26*rendimento_carteira`        |
| 20   | `=FV($D$19; $A27*12; $D$17*-1)`       | `=C27*rendimento_carteira`        |
| 30   | `=FV($D$19; $A28*12; $D$17*-1)`       | `=C28*rendimento_carteira`        |

Cada linha reutiliza a taxa mensal (`D19`) e o aporte (`D17`) e muda apenas o número de meses (`A*12`).

### Divisão do aporte por perfil

A célula de perfil é combinada com o tipo de FII para formar uma **chave** de busca:

```excel
=VLOOKUP($C$32&"-"&B36; Planilha2!$A:$D; 4; FALSE)
```

- `$C$32&"-"&B36` gera chaves como `Conservador-PAPEL` ou `Agressivo-TIJOLO`.
- A tabela de apoio (`Planilha2`) contém as colunas **CHAVE**, **PERFIL**, **TIPO DE FII** e **%**, com a chave montada por `=B3&"-"&C3`.
- O valor retornado (coluna 4) é o percentual do tipo de FII; multiplicado pelo aporte (`$C$33`), gera o **valor a investir** em cada categoria.
- A soma dos valores (`=SUM(D36:D41)`) confere se a distribuição fecha com o aporte total.

#### Tabela de apoio (percentuais por perfil)

| Tipo de FII    | Conservador | Moderado | Agressivo |
| -------------- | ----------- | -------- | --------- |
| PAPEL          | 30%         | 32%      | 50%       |
| TIJOLO         | 50%         | 35%      | 10%       |
| HÍBRIDOS       | 10%         | 8%       | 5%        |
| FOFs           | 10%         | 5%       | 5%        |
| DESENVOLVIMENTO| 0%          | 10%      | 20%       |
| HOTELARIAS     | 0%          | 10%      | 10%       |

## Ferramentas e funções utilizadas

- `VF` / `FV` — valor futuro de uma série de aportes.
- `PROCV` / `VLOOKUP` — busca do percentual por perfil e tipo de FII.
- **Intervalos nomeados** — fórmulas mais legíveis e reutilizáveis.
- **Tabela de apoio** — centraliza os percentuais e facilita a manutenção.
- Formatação condicional e de moeda/percentual para leitura dos resultados.

## Aprendizados

- Traduzir um problema de investimento em um modelo de cálculo no Excel.
- Usar `VF` para projetar patrimônio a partir de aportes periódicos.
- Separar dados de entrada, cálculos e resultados com intervalos nomeados.
- Estruturar tabelas de apoio para consultas com `PROCV`.
- Comparar cenários de longo prazo e entender o efeito dos juros compostos.

## Autor

Projeto desenvolvido como parte da Formação DIO.

- Nome: Alexandre Pereira Santos
- LinkedIn: [Alexandre Santos](https://www.linkedin.com/in/alexendrios/)
- GitHub: [Alexandre Santos](https://github.com/alexendrios)
