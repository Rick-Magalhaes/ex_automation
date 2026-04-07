Automação de Votação

Script em Python para automatizar o preenchimento de planilhas Excel a partir de arquivos PDF organizados em pastas por empresa.

Ele faz:

Leitura dos nomes dos arquivos
Extração de votos
Matching inteligente com nomes da planilha
Preenchimento automático do Excel
Identificação de registros encontrados e não encontrados

Como funciona
Você fornece:
Uma pasta com subpastas (cada uma representando uma empresa)
Um arquivo Excel base
O sistema:
Lê todos os arquivos das pastas
Extrai nome + votos do nome do arquivo
Normaliza os nomes (remove acento, padroniza)
Faz matching inteligente com a planilha
Preenche os votos automaticamente
Gera:
Um novo Excel com sufixo _atualizado

Estrutura esperada
/pasta_base/
   /Empresa1/
      João Silva - A;R;A.pdf
      Maria Souza - A;A;R.pdf

   /Empresa2/
      Pedro Lima - R;A;A.pdf

Regras de Matching (IMPORTANTE)

O sistema usa duas estratégias para casar nomes:

1. Subconjunto de palavras (PRIORIDADE ALTA)

Exemplo:

Arquivo: Alberto ROBERTO
Planilha: ALBERTO ROBERTO ALVES

→ MATCH 

2. Fuzzy Matching (similaridade)
Usa biblioteca rapidfuzz
Threshold configurado: 85%

Problemas que o sistema resolve

✔ Evita duplicações (ex: vários "Flávio")
✔ Evita matches errados (ex: "Marcos Flávio" ≠ "Flávio")
✔ Garante 1 arquivo → 1 linha
✔ Controla nomes já utilizados

Requisitos

Python 3.10+

Instale as dependências:

pip install -r requirements.txt

requirements.txt
openpyxl
rapidfuzz

Como executar
python main.py

Você deverá informar:

Caminho da pasta com arquivos:
Caminho do arquivo Excel:

Saída

O sistema gera:

Novo arquivo Excel:
nome_original_atualizado.xlsx
Coluna de status:
ok → encontrado
fora → não encontrado

Configurações importantes

No código:

FUZZY_THRESHOLD = 85
MIN_PALAVRAS_SUBSET = 2

Você pode ajustar:

Sensibilidade do matching
Rigor do subconjunto

Logs

O sistema exibe:

Arquivos duplicados
Problemas de formatação
Quantidade de matches

Exemplo:

INFO | Colunas de itens detectadas: 5
WARNING | Duplicados encontrados:
WARNING | FLAVIO — 2 arquivos
INFO | Encontrados: 18 | Não encontrados: 0

Executável (.exe)

Se você gerar um .exe com PyInstaller:

pyinstaller --onefile main.py

✔ Pode enviar só o .exe
✔ Não precisa enviar Python nem bibliotecas

Apenas garanta que o usuário tenha:

O Excel
As pastas no formato correto

Boas práticas
Padronize nomes dos arquivos:
NOME COMPLETO - VOTO;VOTO;VOTO.pdf
Evite:
nomes incompletos
caracteres estranhos
formatos diferentes
 Escalabilidade

O sistema foi pensado para:

30.000+ arquivos ✅

Com:

controle de duplicidade
matching eficiente
uso de estruturas otimizadas

Possíveis melhorias futuras
Interface gráfica (GUI)
Leitura direta do conteúdo do PDF
Exportação de relatório de inconsistências
Paralelização para grandes volumes