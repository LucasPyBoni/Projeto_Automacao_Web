## Sobre o Projeto: 

Pesquisa automatizada utilizando selenium. O programa lê a tabela do excel e as duas funções criadas (buscador_Google e buscador_buscape), retornam as pesquisas filtradas, além de transformar em tabela do excel e enviar automaticamente por email ao finalizar.



### Principais bibliotecas:

Selenium, Pandas e win32.com

### Explicação da linha de raciocínio:
 
O navegador é executado, e logo após, a tabela excel é lida. As funções recebem o navegador e as informações da tabela, depois os da tabela são tratados para abranger a busca. Através do selenium o navegador do google shopping ou buscapé é aberto, o find_element(s) retorna as informações para uma lista. Dessa forma, pra cada produto na tabela excel a função retorna seus dados e armazena em um dataframe, que por sua vez é transformado em arquivo excel por meio do pandas.read para enviar por email.
