## Resumo: 
Esse projeto foi desenvolvido com a finalidade de automatizar tarefas de downloads de arquivos em um site que uma empresa utiliza diariamente. A versão final do projeto foi compilada em um arquivo executável (.exe) para o uso.

## Metodologia:
- O principal elemento do código é a biblioteca do **Selenium**, que basicamente fornece métodos de interagir com os elementos do navegador (ex.: button e input);
- Outra bibliteca usada é a **OS** na qual interage com o sistema operacional do computador para renomear, apagar e redirecionar o caminho dos arquivos baixados;
- Biblioteca **time** fora usada para manipulação do código utilizando o horário para executar alguma ação;
- A **webdriver_manager.chrome** é outra biblioteca muito importante para o projeto. Além do Selenium, é necessário uma library de navegador para que o mesmo funcione, e neste caso o navegador utilizado é o Google Chrome;
- **Schedule** nos permite rodar o código de tempos em tempos;
- **from tkinter import messagebox** conforme autoexplicativo, utilizei para que, ao final do código, exiba um pop-up informando a finalização da automatização.
