# 📊 Dashboard de Resultados — Porto Seguro S.A.

Dashboard executivo interativo construído a partir dos dados públicos de Relações com Investidores da Porto Seguro, transformando o release trimestral em uma ferramenta de análise no padrão FP&A.

![Demonstração do dashboard](assets/dashboard-demo.gif)

> ⚠️ **Projeto independente e educacional.** Utiliza exclusivamente dados públicos divulgados pela companhia em seu site de RI. **Não constitui recomendação de investimento** e não possui qualquer vínculo com a Porto Seguro S.A.

---

## 🎯 Objetivo

Releases trimestrais são excelentes para investidores, mas pouco práticos para quem precisa analisar rapidamente a evolução operacional de um negócio. Este projeto converte a planilha de resultados da companhia em um **dashboard HTML de arquivo único** — sem backend, sem instalação — com a visão que um analista de FP&A realmente usa: DRE gerencial com variações, indicadores por vertical, análise de crescimento e insights automáticos.

## 🖥️ Demonstração

**[▶ Acessar o dashboard ao vivo](https://andremleite.github.io/FPA-dashboard-porto-seguro/)**

*(prints em `/screenshots`)*

## ⚙️ Funcionalidades

| Área | Recursos |
|---|---|
| **Resumo Executivo** | 10 KPIs com variação YoY/QoQ e sparklines · Destaques do trimestre · DRE Gerencial com Δ QoQ/YoY · Indicador em Foco dinâmico |
| **Segmentos** | Cards por vertical (Seguro, Saúde, Bank, Serviço) com mini-KPIs e gráficos combinados |
| **Análise Financeira** | Dispersão Receita×Lucro · Bullet charts de margens · Evolução patrimonial · Waterfall de estrutura de capital |
| **Crescimento** | Heatmap QoQ · Barras divergentes YoY/QoQ · Ranking CAGR — comutável por indicador (Receita, Lucro, ROAE, índices em Δ p.p.) |
| **Mercado** | LPA, VPA, ROAE e política de proventos |
| **Insights** | 19 insights gerados automaticamente dos dados, com tags por vertical |
| **Interatividade** | Filtros globais (período, segmento, indicador) com *cross-filtering* em todas as abas · Comparação de dois trimestres lado a lado · Zoom de gráficos em modo apresentação com valores plotados · Drill-down |
| **Exportação** | PNG e PDF (com valores nos gráficos) · Excel com 5 abas de dados · Botão "123" para exibir valores em prints manuais |
| **Idioma** | Interface bilíngue PT/EN com um clique — inclusive formatação numérica e de trimestre |
| **Experiência** | Tema claro/escuro · Identidade visual Porto · Responsivo · Navegação por teclado (← →) |

## 🛠️ Tecnologias

- **Front-end:** HTML5, CSS3, JavaScript ES6+ (arquivo único, `index.html`)
- **Visualização:** [Chart.js 4](https://www.chartjs.org/) · [ApexCharts](https://apexcharts.com/) (waterfall)
- **UI:** Tailwind CSS (Play CDN) · Font Awesome · Plus Jakarta Sans / Inter
- **Exportação:** html2canvas · jsPDF · ExcelJS
- **Pipeline de dados:** Python (`openpyxl`) — extração da planilha de RI para o JSON embutido no HTML

## 🔄 Como atualizar com um novo trimestre

A arquitetura separa **dados** de **lógica**: toda a base fica em um bloco JSON embutido no `index.html`. Para atualizar:

**Opção 1 — Google Colab (sem instalar nada):**
Abrir o notebook [`notebooks/atualizar_dashboard_colab.ipynb`](notebooks/atualizar_dashboard_colab.ipynb) no [Google Colab](https://colab.research.google.com/), enviar a planilha nova de resultados + o `update_dashboard.py` + o `index.html` atual, e rodar as células em ordem. O notebook devolve o dashboard já atualizado para download.

**Opção 2 — localmente, via terminal:**
```bash
pip install openpyxl
python update_dashboard.py Planilha-3T26.xlsx --template index.html --out index.html
```

Em ambos os casos, o script:
1. Localiza cada linha da planilha pelo **texto do rótulo** (ex.: "Lucro Líquido"), não pela posição — resiliente a pequenas mudanças de layout entre trimestres
2. Extrai a série completa, de 2023 até o trimestre mais recente
3. Roda uma **validação de sanidade** (receita, lucro e indicadores dentro de intervalos plausíveis) antes de publicar
4. Gera o JSON e injeta no template HTML — layout, filtros, gráficos e insights automáticos permanecem os mesmos; só os dados mudam

> A planilha original da companhia não é redistribuída neste repositório — baixe-a diretamente no [Kit do Investidor da Porto](https://ri.portoseguro.com.br/).

## 📁 Estrutura do repositório

```
FPA-dashboard-porto-seguro/
│
├── index.html               # o dashboard completo (arquivo único, servido pelo GitHub Pages)
├── update_dashboard.py       # script de atualização (planilha → JSON → HTML)
├── notebooks/
│   └── atualizar_dashboard_colab.ipynb   # mesmo processo, rodável no navegador via Colab
├── data/
│   └── dashboard.json        # dados extraídos na última atualização (opcional, para transparência)
├── assets/
├── screenshots/
├── README.md
└── LICENSE
```

## 🗺️ Roadmap

- **v1.0** — Dashboard executivo 1T26: KPIs, DRE, verticais, crescimento, insights, exports ✅
- **v1.1** — Comparação de trimestres lado a lado · filtro de indicador em todas as abas · interface bilíngue PT/EN ✅
- **v1.2** — Atualização 2T26/1S26 via pipeline Python + notebook Colab documentado ✅
- **v2.0** — Comparação com pares do setor (seguradoras listadas) · múltiplos de mercado

## 📄 Licença

MIT — livre para estudar, adaptar e reutilizar, mantendo a atribuição.

---

**Criado por André Leite** · FP&A & Investor Relations
Feedbacks e sugestões são bem-vindos!
