# 📦 Automação de Cotação e Controle de Estoque com Excel VBA
🗓️ Data de desenvolvimento em: Outubro de 2024

Este projeto consiste em duas rotinas desenvolvidas em **Excel VBA** para automatizar o processo de **cotação de pedidos** e **atualização de estoque** a partir de planilhas extraídas do sistema Linear.

---

## ⚙️ Funcionalidades

🧠 Diferenciais de Engenharia
Navegação Relativa (Offset): O código não utiliza coordenadas fixas de colunas. Ele calcula a posição final do relatório dinamicamente, garantindo que a automação continue funcionando mesmo se o sistema Linear adicionar novas colunas no futuro.
Sanitização Automática (ETL): Inclui uma rotina de limpeza que remove cabeçalhos inúteis e converte strings em valores numéricos, garantindo a integridade dos cálculos matemáticos.
Persistência de Dados (Snapshots): Utiliza a técnica de converter fórmulas em valores estáticos após o processamento, registrando o histórico do estoque com carimbo de data sem risco de alteração posterior por recalculo automático.

### 🧾 Cotação de Pedidos (`Sub Cotacao_Fernando`)
- Abre automaticamente planilhas externas contendo pedidos.
- Localiza e extrai informações de cada pedido (produto, código, estoque).
- Alimenta uma planilha central de cotações com os dados organizados.
- Aplica **formatação condicional** para destacar pedidos vencidos ou em risco de vencimento.

### 📊 Controle de Estoque (`Sub Estoque_Fernando`)
- Abre o relatório de estoque (`00dado.xls`) exportado do sistema Linear.
- Remove linhas desnecessárias e converte dados textuais em numéricos.
- Realiza buscas automáticas (`PROCV`) para localizar o estoque atual de cada produto.
- Atualiza as colunas de estoque e pedido, com a data do dia.
- Aplica bordas, alinhamento e cores para padronização visual.

---

## ⏱️ Impacto

- Tempo médio anterior: **~20 minutos por cotação**
- Tempo atual com automação: **< 2 minutos**
- Redução de tempo: **+90% de eficiência**
- Maior confiabilidade e padronização nas análises

---

## 💡 Tecnologias e técnicas utilizadas

- **Excel VBA**
- Manipulação de múltiplas planilhas e arquivos externos
- `PROCV`, `SEERRO`, `HOJE`, e outras fórmulas automatizadas
- Limpeza de dados e aplicação de formatos
- Automação de tarefas repetitivas com laços (`Do While`, `If`)

---

## ✅ Competências demonstradas

- Automação de rotinas administrativas
- Organização de dados para tomada de decisão
- Estruturação de macros eficientes e reutilizáveis
- Redução de erros manuais em processos críticos
- Pensamento lógico e foco em produtividade

---

## 📌 Projeto de uso interno e educativo.  
**Desenvolvido por Caio Cesar de Albuquerque**  
📫 [caioalbuquerquedev@gmail.com](mailto:caioalbuquerquedev@gmail.com)  
🔗 [LinkedIn](https://www.linkedin.com/in/caio-cesar-for-hire) | [GitHub](https://github.com/Caio-Cesa)


