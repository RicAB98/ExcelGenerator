# Instruções

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Esta ferramenta foi criada com o intuito de ajudar na criação de grelha de avaliação em ficheiro Excel. Apesar da aplicação ser capaz de gerar grelhas funcionais, as opções são bastante limitadas e por essa razão não permite a criação de grelhas mais complexas. Todavia, é possível utilizar esta ferramenta para criar um versão simplicada da tabela, desproteger as folhas e fazer as alterações necessárias para acelerar o processo.

## Botões
* **Carregar ficheiro**: sempre que um ficheiro Excel é gerado a partir da ferramenta, é também criado um ficheiro em formato JSON, que consiste nas configurações da grelha. Carregar este ficheiro, copia as configurações que lá se encontram para a aplicação. Isto foi feito para efeitos de desenvolvimento da ferramenta, para facilmente poder replicar ficheiros que tenham sido gerados com erros. No entanto, podem também ser utilizado para partilhar configurações e evitar ter que o fazer manualmente.  
* **Adicionar secção** e **Remover secção**
* **Selecionar**: abre uma janela para escolher a cor da secção correspondente
* **Gerar**: se as configurações inseridas forem validadas, gera os ficheiros Excel e Json para a pasta "FicheirosGerados". 

## Campos
* **Nome do ficheiro**: por exemplo, assumindo que o nome inserido foi Grelha, o ficheiro será criado como "Grelha.xlsx" (sendo .xlsl o sufixo/extensão dos ficheiros Excel)
* **Nº alunos**: número de linhas que serão criadas em cada Período
* **Nº períodos**: esta opção foi criada devida a algumas escolas passarem para dois períodos em vez dos tradionais três.
* **Secções**: área onde serão indicadas as diferentes áreas de avaliação (ex: Atitudes e Valores, Testes, Trabalhos ...). Células com fundo cinzento têm o seu valor calculado e por isso não é permitida a sua edição manual.
    *  **Nome da secção**
    *  **Nº subsecções**: número de componentes de cada secção (ex: Secção "Testes" pode ser composta por três subsecões - "Teste 1", "Teste 2", "Teste 3"). É dado o mesmo peso a cada uma das subsecções.
    *  **Percentagem**: peso da secção para a nota final do período
    *  **Cor de fundo** e **Cor de texto**: opções para customizar o cabeçalho de cada secção para quem preferir grelhas policromáticas
    ![Imagem](https://i.imgur.com/KmDCvvd.png)
* **Avaliação**: opções para customizar o cabeçalho da secção de "Avaliação" para quem preferir grelhas policromáticas. Esta secção consiste em 4 colunas:
    * **Média do perído**: cálculo da nota tendo em conta as diferentes secções do período
    * **Média final**: média do período atual com os anteriores
    * **Média final**: auto-avaliação do aluno
    * **Nível a atribuir**: nota final a atribuir ao aluno
    ![Imagem](https://i.imgur.com/dWx4zl4.png)

## Zona de feedback
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;À esquerda do botão "Gerar", existe uma zona onde ocasionalmente aparecerão mensagens com informaçõs pertinentes, nomeadamente quando algum erro é encontrado ou quando a geração de ficheiros terminar.

![Imagem](https://i.imgur.com/NBJT6PO.png)