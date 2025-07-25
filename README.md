
🤖 Automação de Extração de Dados de Faturas PDF
Este projeto contém um script Python para automatizar o processo de leitura e extração de dados financeiros de faturas em formato PDF. O script foi desenvolvido para ser robusto e flexível, lidando com variações de layout entre os documentos.

O robô lê múltiplos arquivos PDF de uma pasta, extrai informações chave (como número de telefone, valor bruto, retenção e valor a pagar), aplica regras de negócio para calcular valores ausentes e, por fim, gera uma planilha Excel (.xlsx) consolidada e formatada com os resultados.

✨ Funcionalidades Principais
Leitura em Lote: Processa todos os arquivos .pdf encontrados em um diretório específico.

Extração Inteligente (Verificação Dupla): Utiliza uma abordagem híbrida para garantir a extração dos dados:

Método por Coordenadas (GPS): Tenta primeiro localizar os dados com alta precisão, baseado na posição dos textos na página.

Método por Regex (Mapa): Caso o primeiro método falhe, um plano B é ativado para fazer uma busca por padrões de texto, garantindo a captura dos dados mesmo em PDFs complexos.

Lógica de Autocorreção: Se um valor (como o Valor Bruto) não for encontrado, o script o calcula automaticamente com base nos outros valores extraídos, seguindo a fórmula Bruto = Líquido + Retenção.

Flexibilidade: O script é capaz de encontrar dados mesmo com pequenas variações nos rótulos (ex: busca por "VALOR DA RETENCAO IMPOSTOS" e também por "RETENCOES").

Relatório em Excel: Gera um relatório consolidado e profissional em formato .xlsx.

Formatação Automática: As colunas de valores na planilha são automaticamente formatadas como moeda (R$).

Totalização: Adiciona uma linha de TOTAL ao final do relatório, somando as colunas de valores.

🛠️ Tecnologias Utilizadas
Python 3

PyMuPDF (fitz): Para a leitura e extração de dados dos arquivos PDF.

Pandas: Para a manipulação dos dados, criação da tabela e exportação para o Excel.

Openpyxl: Como motor para o Pandas escrever arquivos no formato .xlsx.

⚙️ Pré-requisitos
Antes de começar, garanta que você tenha o Python 3 instalado em seu sistema. Você pode baixá-lo em python.org.

🚀 Instalação e Configuração
Siga os passos abaixo para configurar o projeto em sua máquina local.

Clone o repositório:

Bash

git clone https://github.com/seu-usuario/seu-repositorio.git
(Substitua pelo link do seu repositório no GitHub)

Navegue até a pasta do projeto:

Bash

cd seu-repositorio
Instale as dependências:
Abra o seu terminal e execute o seguinte comando para instalar todas as bibliotecas necessárias de uma só vez:

Bash

pip install PyMuPDF pandas openpyxl
▶️ Como Usar
Com o ambiente configurado, siga os passos abaixo para executar a automação:

Crie a pasta de PDFs: Dentro da pasta principal do projeto, crie uma nova pasta e nomeie-a exatamente como boletos_pdf.

Adicione os arquivos: Coloque todos os seus arquivos de fatura em formato .pdf dentro da pasta boletos_pdf.

Execute o script: Abra o seu terminal na pasta principal do projeto e rode o seguinte comando:

Bash

python automacao_boletos.py
Verifique o resultado: O script irá processar cada arquivo e, ao final, criará uma nova planilha Excel chamada Relatorio_Final_Corrigido.xlsx na pasta do projeto. Esta planilha conterá todos os dados extraídos, calculados e formatados.

🧠 Lógica do Código
O script automacao_boletos.py é o coração do projeto. Ele contém funções auxiliares para tarefas como limpar valores monetários e uma função principal que utiliza a estratégia de "Verificação Dupla" para garantir a extração dos dados. Ao final, ele usa a biblioteca Pandas para gerar um relatório profissional e consolidado.


