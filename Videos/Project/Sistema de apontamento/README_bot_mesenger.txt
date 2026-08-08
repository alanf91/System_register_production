ALGORITMO: EXCEL -> WHATSAPP

1. ARQUIVOS
Coloque estes arquivos na mesma pasta:
- notificador_whatsapp_excel.py
- .env
- controle_transferencias.xlsx

2. ABA E COLUNAS
A planilha deve possuir uma aba chamada:
Dados

Cabeçalhos obrigatórios na primeira linha:
ID
FABRICA
TELEFONE
ITEM
QUANTIDADE
DATA_ENVIO
DATA_RECEBIMENTO

O programa cria automaticamente:
AVISO_ENVIO
AVISO_RECEBIMENTO
DATA_ULTIMO_AVISO
RETORNO_WHATSAPP

3. INSTALAÇÃO
Abra o terminal na pasta e execute:

pip install openpyxl requests python-dotenv

4. CONFIGURAÇÃO
Renomeie:
.env.exemplo
para:
.env

Depois preencha o token e o Phone Number ID fornecidos pela Meta.

5. TEMPLATE PARA PRODUÇÃO
Crie e aprove na Meta um template chamado:
atualizacao_movimentacao_fabrica

Sugestão de corpo:

Atualização de movimentação da fábrica {{1}}:

{{2}}

Idioma:
Português (Brasil) - pt_BR

6. EXECUÇÃO
No terminal:

python notificador_whatsapp_excel.py

7. FUNCIONAMENTO
- Se DATA_ENVIO estiver preenchida e AVISO_ENVIO não for SIM,
  o evento de envio será comunicado.
- Se DATA_RECEBIMENTO estiver preenchida e
  AVISO_RECEBIMENTO não for SIM, o recebimento será comunicado.
- Os eventos são agrupados por fábrica e telefone.
- Após sucesso, o programa grava SIM na coluna correspondente.
- Antes de alterar a planilha, o programa cria uma cópia na pasta backup.

8. AGENDAMENTO
No Windows, use o Agendador de Tarefas para executar o arquivo Python
em horários definidos, por exemplo, a cada 15 minutos ou uma vez por hora.
