*Extração Multi-Thread ERP v1.2
Este projeto é uma ferramenta de automação de alta performance desenvolvida em Python, utilizando Selenium e Multi-threading para extração de dados de produtos em sistemas ERP via web.*

O sistema foi otimizado para processar grandes volumes de dados (bases de +40k itens), alcançando marcas de 12.5 itens por segundo (aprox. 100 itens a cada 8 segundos).

*Funcionalidades
Processamento Paralelo: Suporte a múltiplas instâncias simultâneas do Chrome (Threads).

Configuração Dinâmica: Escolha a coluna e a linha inicial do Excel diretamente pela interface.

Filtro Inteligente: Limpeza automática de termos indesejados nas referências.

Tratamento de Dados: Identifica campos vazios no ERP e registra como "vazio" no log para facilitar a busca (Ctrl+F).

Interface Intuitiva: Desenvolvida em Tkinter com monitoramento de progresso em tempo real.

*Tecnologias Utilizadas
Python 3.x

Pandas: Manipulação de grandes planilhas Excel.

Selenium: Automação de navegação web.

ThreadPoolExecutor: Gerenciamento de concorrência.

Tkinter: Interface gráfica de usuário (GUI).

*Pré-requisitos
Para rodar o código fonte, você precisará instalar as dependências:

pip install pandas selenium webdriver-manager openpyxl

*Estrutura do Excel
O programa espera uma planilha .xlsx onde:

Coluna de Código: Definida na interface (ex: Coluna C).

Coluna de Referência: O sistema lê automaticamente a coluna imediatamente à direita da definida para o código.

*Como gerar o Executável
Para criar o instalador único (.exe), utilize o PyInstaller:

pyinstaller --onefile --noconsole --name "Automacao_ERP_v1.2" main.py

*Configurações Recomendadas
Para bases de dados acima de 40.000 itens, recomenda-se:

Threads: Entre 8 e 12 (dependendo da sua memória RAM).

Linha de Início: 16 (ajustável conforme o cabeçalho da sua planilha).

Desenvolvido por Marco 
