# Excel-MonteCarlo-VBA
Este é um código em VBA que irá permitir que rode Simulações Monte Carlo (Distribuição PERT) direto no seu Excel, facilitando a entrada em estimativas e simulações de cenários.

#  Simulação de Monte Carlo no Excel (VBA)

O modelo utiliza a **Simulação de Monte Carlo** baseada na **Distribuição PERT** para projetar cenários de custos, prazos ou retornos financeiros, transformando estimativas estáticas em análises probabilísticas dinâmicas. Ideal para gerentes de projeto, analistas financeiros e engenheiros.

##  Destaques do Modelo

* **Setup Automatizado:** Com um clique, o VBA formata a interface da planilha e cria botões interativos modernos para o usuário.
* **Motor Matemático Robusto:** Utiliza o algoritmo de aceitação-rejeição de Jöhnk para gerar números pseudoaleatórios que respeitam a curva da Distribuição Beta/PERT.
* **Flexibilidade de Dados:** 
  * Permite usar um **Histórico Real** (o modelo calcula automaticamente a Mediana, Mínimo e Máximo).
  * Permite usar **Estimativa de 3 Pontos** (Otimista, Mais Provável e Pessimista).
* **Análise de Alvo:** O usuário insere um "Orçamento/Prazo Alvo" e o algoritmo calcula a exata probabilidade de sucesso daquele cenário.
* **Gráficos Executivos (Automáticos):** Gera nativamente uma **Curva S (Probabilidade Acumulada)** e um **Histograma de Densidade**.

##  Como Baixar e Usar

Como este é um projeto em Excel com macros, você precisa baixar o arquivo para a sua máquina:

1. Neste repositório, clique no arquivo `Monte_Carlo_Simulation_VBA.xlsm`.
2. Clique no botão de **Download** (ícone de seta apontando para baixo ⬇️) localizado no lado direito da tela.
3. Abra o arquivo no seu Microsoft Excel.
4. O Excel mostrará um aviso amarelo de segurança no topo ("As macros foram desabilitadas"). Clique em **Habilitar Conteúdo**.

### Passo a Passo de Uso:
1. Vá na aba **Exibir > Macros** e rode a macro `SetupTemplate` para o sistema desenhar a interface e gerar o botão de simulação. (Você só precisa fazer isso uma vez).
2. Preencha as células verdes com suas estimativas.
3. Clique no botão verde **▶ Rodar Simulação**.
4. Os percentis (P10, P50, P90) e os gráficos de Curva S e Densidade aparecerão imediatamente.

##  Requisitos e Compatibilidade
* **Microsoft Excel (Windows):** O código VBA foi testado e otimizado para as versões desktop do Excel.
* **Versão Mínima para Gráficos Completos:** Excel 2016 ou superior (Microsoft 365). O gráfico de *Histograma nativo* foi introduzido no Excel 2016. Se você usar uma versão mais antiga (ex: Excel 2013), o motor matemático e a Curva S funcionarão perfeitamente, mas o script pulará a geração do Histograma para evitar erros.

##  Visualizando o Código (VBA)
Se você é desenvolvedor ou analista e quer estudar a lógica por trás da simulação:
1. Com a planilha aberta, pressione `Alt + F11` para abrir o Editor do VBA.
2. Na árvore de projetos à esquerda, abra a pasta `Módulos` e dê um duplo clique em `Módulo 1`. Todo o código estará comentado e estruturado lá.

##  Contribuições e Contato
Sinta-se livre para fazer um *fork* deste projeto, sugerir melhorias na eficiência do laço de repetição (`For Next`) ou adicionar cálculos de novas distribuições (Normal, Lognormal, Triangular). 

Se você gostou deste projeto, me conecte no LinkedIn! (https://www.linkedin.com/in/danielnacarat/)
