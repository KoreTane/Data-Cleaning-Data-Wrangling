Resumo do Projeto:
Este projeto visa gerar um relatório DRE (Demonstração do Resultado do Exercício) em Power BI para apresentação aos diretores de uma empresa que presta serviços pessoais com sede no Brasil e nos EUA. O objetivo é criar um relatório simples e fácil de entender para pessoas leigas e experientes em negócios.

Objetivos:
Criar um relatório DRE fácil de entender para diretores da empresa
Utilizar Python para tratamento de dados e preparação para modelos de ML
Integrar com Power BI para visualização de dados

Estrutura do Projeto:
O projeto é dividido em duas bases de dados: Base BR e Base USA, cada uma com seus respectivos planos de contas.

Tratamento de Dados:
O tratamento de dados é feito em Python, com o objetivo de preparar a base para possíveis modelos de Machine Learning (ML) no futuro. O arquivo .py atualiza o relatório em Power BI.

Tecnologias Utilizadas:
Python
Power BI

Entrega do Projeto:
O projeto foi entregue com a base de dados tratada em Python e toda a engenharia de recursos necessária. Ao executar o arquivo .py, o relatório é atualizado automaticamente.

Status do Projeto:
Concluído.

OBS: Estou indo além do projeto inicial e também calculando a DRE em Python. O plano de contas está pronto e se conecta a um relatório DRE no Excel por meio de VBA, utilizando consultas SQL para recuperar os dados. O relatório é simplificado e inclui análise comparativa temporal.


Dependências e Instalação:
Para executar este projeto, você precisará instalar as seguintes dependências:
  
  pandas-profiling versão 3.3.0: pip install pandas-profiling==3.3.0
  
  xlsxwriter: pip install xlsxwriter
  
  pandas versão 2.0.3: pip install pandas==2.0.3

Importação de Bibliotecas:
O projeto utiliza as seguintes bibliotecas:
  
  pandas como pd: para manipulação e análise de dados
  
  numpy como np: para computações numéricas
  
  babel.dates: para lidar com datas e horários em aplicações Python
