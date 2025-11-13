# pyCatapulta — Frontend para openMSX

pyCatapulta é um frontend para o emulador `openMSX` (Windows) que permite iniciar o emulador com opções selecionadas e, opcionalmente, registrar configurações por execução em um banco SQLite. Projeto de longo prazo em desenvolvimento.

## Funcionalidades
- Interface gráfica (CTk) para iniciar `openmsx.exe`.
- Exibição do PID do processo do `openmsx.exe` (tenta detectar o PID "real" usando `psutil` quando disponível).
- Mostra o caminho do socket gerado pelo openMSX e permite abrir o local no Explorer.
- Armazena configurações em SQLite para cada execução, se o usuário desejar.
- Janela de configuração para apontar o diretório contendo `openmsx.exe`.

## Tecnologias
- Python (Windows)
- GUI: `customtkinter`
- Banco de dados: `sqlite3` (arquivo local)
- Opcional: `psutil` (melhora detecção do PID)
- Ferramentas: `pip`, `python`

## Requisitos
- Windows 10/11
- Python 3.8+
- Pacotes Python:
  - `customtkinter`
  - `psutil` (opcional, recomendado)
- `openmsx.exe` disponível em um diretório configurado pela aplicação

## Instalação
1. Clone o repositório:
   - `git clone git@github.com:wilsonpilon/pyCatapulta.git`
2. Entre na pasta do projeto:
   - `cd pyCatapulta`
3. Instale dependências:
   - `pip install customtkinter`
   - `pip install psutil` (opcional)
4. Execute:
   - `python main.py`

## Uso
- Ao abrir `main.py`, configurar o diretório do `openMSX` via `Configuração`.
- Usar `Iniciar openMSX` para executar o emulador.
- O PID exibido é o PID detectado; quando `psutil` não estiver presente, o PID inicial do processo lançado será usado.
- O caminho do socket é exibido na interface; use `Ver Socket` para abrir a pasta no Explorer (se existir).
- Configurações e últimos dados ficam em `app_config.json` e no banco `app_data.db` (padrão).

## Arquivos importantes
- `main.py` — aplicação principal (GUI e lógica).
- `app_config.json` — arquivo de configuração da aplicação.
- `app_data.db` — banco SQLite com registros e configurações.

## Observações
- Esta aplicação foi projetada exclusivamente para Windows.
- Projeto em desenvolvimento contínuo; features e UX serão aprimorados ao longo do tempo.
- O comportamento de detecção de PID e criação de sockets depende do `openMSX` e do ambiente onde ele é executado.

## Contribuição
Contribuições são bem-vindas. Abra issues para bugs/ideias e pull requests com mudanças bem descritas.

## Licença
Licença a ser definida pelo autor do repositório.
