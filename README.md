<div align="center">

# 📈 Portfólio & Controle de Investimentos Executivo

  <p align="center">
    Uma solução robusta em Microsoft Excel para gestão de patrimônio, análise de cotações, controle de proventos e rebalanceamento dinâmico de carteira.
  </p>

[![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](#)
[![License](https://img.shields.io/badge/License-MIT-blue.svg?style=for-the-badge)](#)
[![Status](https://img.shields.io/badge/Status-Concluído-success?style=for-the-badge)](#)

</div>

---

## 📌 Visão Geral

O **Portfólio de Investimentos Executivo** foi desenvolvido para centralizar a gestão de ativos financeiros multi-classe (*Ações, FIIs, Renda Fixa e Criptomoedas*). A planilha automatiza cálculos de rentabilidade, preço médio, alocação percentual e projeção de proventos, eliminando o trabalho manual e oferecendo suporte direto para tomada de decisão.

---

## 🚀 Funcionalidades Principais

- **📊 Dashboard Executivo:** Painel visual com KPIs em tempo real, cards informativos e gráficos de alocação de patrimônio.
- **💼 Posição Consolidada:** Cálculo automático de Preço Médio (PM), Valor Atualizado, Lucro/Prejuízo nominal e percentual.
- **🎯 Análise de Preço Teto:** Sistema condicional que indica se o ativo está no ponto de compra (`Abaixo do Teto`) ou aguardando correção (`Acima do Teto`).
- **💰 Gestão de Proventos & Renda Passiva:** Controle de Dividendos, JCP e Rendimentos de FIIs, integrado ao cálculo de **Yield on Cost (YoC)**.
- **⚖️ Rebalanceamento Automático:** Definição de metas por categoria e sugestão inteligente de aportes em R$ para manutenção de risco.
- **📝 Histórico de Movimentações:** Log parametrizado para registrar compras, vendas e custos operacionais (taxas/corretagem).

---

## 🧱 Estrutura das Abas

| Ícone | Aba | Descrição Técnica |
| :---: | :--- | :--- |
| 🖥️ | **`Dashboard`** | Visão macro do patrimônio, gráficos interativos de alocação e indicadores de desempenho. |
| 📈 | **`Consolidado`** | Tabela dinâmica de custódia conectada aos históricos por fórmulas `SOMASES` e `SUMIF`. |
| ⚖️ | **`Rebalanceamento`** | Matriz de alocação com cálculo de variação entre a carteira real e a meta desejada. |
| 💵 | **`Proventos`** | Módulo de recebíveis agrupado por ativo, data com e data de pagamento. |
| 🔄 | **`Movimentações`** | Livro Razão de ordens de compra e venda para alimentação dos motores de cálculo. |

---

## 📐 Fórmulas & Motores Utilizados

A planilha utiliza encadeamento de funções nativas para manter alta performance sem sobrecarregar o processamento:

```excel
# Consolidação de Quantidade Total em Custódia
=SOMASES(Movimentações!E:E; Movimentações!C:C; A3; Movimentações!B:B; "Compra") - SOMASES(Movimentações!E:E; Movimentações!C:C; A3; Movimentações!B:B; "Venda")

# Cálculo Automático de Total Recebido em Proventos
=SUMIF(Proventos!D:D; A3; Proventos!G:G)

# Validação Condicional do Preço Teto
=SE(F3 <= M3; "Abaixo do Teto"; "Acima do Teto")
