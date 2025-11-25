# conferidor-documentos-excel

📄 Conferidor Automático de Documentos em Planilhas Excel

Ferramenta automatizada para conferência, validação e indexação de documentos físicos utilizando planilhas Excel como fonte de referência.



#Visão Geral do projeto:

Este projeto foi desenvolvido para automatizar a conferência de documentos físicos, verificando se cada documento está ou não registrado em uma planilha Excel.

A ferramenta lê automaticamente dados estruturados de duas formas:

Linha a linha → Um documento por célula

Multi-célula → Vários documentos dentro da mesma célula, separados por vírgula

Ela identifica, normaliza e compara números digitados pelo usuário com os documentos registrados na planilha — marcando os encontrados diretamente no arquivo e gerando relatórios detalhados.



#Funcionalidades:

Inteligência na leitura e organização

> Detecta automaticamente o tipo de planilha:

➤ linha_a_linha

➤ multi_celula


> Normaliza documentos removendo:

Prefixos como CTE, CTE RODOVIARIO:

Letras

Espaços, hífens e caracteres especiais

Sufixos como /1


> Busca inteligente

Identifica documentos mesmo em células com vários valores

Marca visualmente apenas o documento encontrado:

Antes: 1642737/1, 1642800/1  
Depois: [✅ 1642737], 1642800


> Marcação automática na planilha

Adiciona indicador [✅ documento] somente ao item conferido

Não altera os outros dados da célula

Evita duplicações

Detecta se um documento já havia sido conferido


> Relatório final

Ao encerrar, exibe:

Quantidade de documentos conferidos

Quantidade não encontrados

Quantos não estavam presentes fisicamente (não marcados na planilha)

Quantos foram digitados repetidos

Mostra quais foram encontrados / não encontrados


> Backups + Logs

Cria backup automático do arquivo original

Gera arquivo de LOG contendo:

Data / hora

Documento digitado

Status: ENCONTRADO / NÃO ENCONTRADO



> Segurança

Nunca sobrescreve o arquivo original

Salva em um novo arquivo com sufixo _conferido.xlsx



#Tecnologias utilizadas no projeto:

Tecnologia	Função
Python 3.13	Núcleo da aplicação
openpyxl	Manipulação de planilhas Excel
tkinter	Interface para seleção de arquivos e pastas
Regex (re)	Normalização e limpeza dos documentos
pathlib	Manipulação moderna de caminhos
shutil	Criação de backups


#Como usar o programa?

> 1. Execute o script Python

No terminal:

python conferidor_final_3.py


> 2. Escolha a planilha Excel

A ferramenta abrirá uma janela pedindo o arquivo .xlsx.

> 3. Escolha a pasta onde ficarão:

backups

logs

arquivo final conferido

> 4. Digite os documentos um por vez

Exemplos válidos:

123456
1692919
1642737/1


> 5. Para encerrar, digite:

fim


> Exemplo de saída do relatório final
📋 RELATÓRIO FINAL:
✔️ Encontrados: 42
❌ Não encontrados: 7
🔁 Repetidos ignorados: 3
📌 Não marcados na planilha (sem físico): 5

💾 Planilha salva como: controle_conferido.xlsx
📝 Log atualizado com sucesso.



#Estrutura do Projeto
📁 conferir-documentos-excel/
│
├── conferir_documentos.py       # Código principal
├── README.md                    # Documentação do projeto
├── LICENSE                      # Licença MIT
└── .gitignore                   # Arquivos ignorados pelo Git



#Motivação:

Este projeto foi criado para resolver um problema comum em ambientes logísticos e administrativos:
conferir centenas ou milhares de documentos físicos usando Excel como referência.

O processo manual é lento, propenso a erros e dificulta auditorias.
Este sistema automatiza totalmente a conferência, garantindo:

Velocidade

Confiabilidade

Rastreabilidade

Segurança da informação

Organização clara do resultado



#Contribuições:

Pull Requests são bem-vindos!
Sinta-se livre para contribuir com melhorias, refatorações ou novas funcionalidades.



#Licença:

Este projeto está licenciado sob a MIT License – permitindo uso comercial, modificação e distribuição.

