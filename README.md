# pyCatapulta — Frontend para openMSX

pyCatapulta é um frontend para o emulador `openMSX` (Windows) que permite iniciar o emulador com opções selecionadas e, opcionalmente, registrar configurações por execução em um banco SQLite. Projeto em desenvolvimento.

## Funcionalidades
- Interface gráfica (CTk) para iniciar `openmsx.exe`.
- Exibição do PID do processo do `openmsx.exe` (tenta detectar o PID "real" usando `psutil` quando disponível).
- Mostra o caminho do socket gerado pelo openMSX e permite abrir o local no Explorer.
- Armazena configurações em SQLite para cada execução, se o usuário desejar.
- Janela de configuração para apontar o diretório contendo `openmsx.exe`.
- Integração planejada para buscar/baixar programas sob demanda (ex.: File‑Hunter).

## Tecnologias
- Python (Windows)
- GUI: `customtkinter`
- Banco de dados: `sqlite3` (arquivo local)
- Opcional: `psutil` (melhora detecção do PID)
- HTTP (planejado): `requests` ou `aiohttp`
- Ferramentas: `pip`, `python`

## Requisitos
- Windows 10/11
- Python 3.8+ (testado preferencialmente em 3.10/3.11)
- `openmsx.exe` disponível em um diretório configurado pela aplicação
- Pacotes Python (exemplos):
  - `customtkinter`
  - `psutil` (opcional, recomendado)
  - `requests` (recomendado para integração de download)

## Instalação
1. Clone o repositório:
   - `git clone git@github.com:wilsonpilon/pyCatapulta.git`
2. Entre na pasta do projeto:
   - `cd pyCatapulta`
3. Instale dependências:
   - Crie/atualize `requirements.txt` (ex.: `customtkinter`, `psutil`, `requests`) e depois:
   - `pip install -r requirements.txt`
4. Execute:
   - `python main.py`

## Uso
- Ao abrir `main.py`, configurar o diretório do `openMSX` via `Configuração`.
- Usar `Iniciar openMSX` para executar o emulador.
- O PID exibido é o PID detectado; quando `psutil` não estiver presente, o PID inicial do processo lançado será usado.
- O caminho do socket é exibido na interface; use `Ver Socket` para abrir a pasta no Explorer (se existir).
- Configurações e últimos dados ficam em `app_config.json` e no banco `app_data.db` (padrão).

## Integração com File‑Hunter (planejada)
- Objetivo: permitir busca e download sob demanda de programas para usar com `openMSX`.
- Verificar API/documentação de `https://download.file-hunter.com/` (endpoints, autenticação, limites e termos).
- Implementação recomendada:
  - Cliente HTTP simples (ex.: `requests`) com métodos `search` e `download`.
  - Diretório de downloads local: `APP_DIR / downloads`.
  - UI: campo de busca, listagem de resultados, botão de download e indicador de progresso.
  - Após download: validar (checksum/tamanho) e oferecer mover/copiar para `share/extensions` ou `share/roms` conforme o tipo do arquivo.
- Observações legais: respeitar termos de uso/licenciamento do site antes de automatizar downloads.

## Configurações persistidas
Chaves salvas pelo app (exemplos):
- `openmsx_dir` — caminho para `openmsx.exe`
- `openmsx_machine` — máquina selecionada
- `openmsx_extensions` — lista JSON de extensões selecionadas
- `openmsx_pid` — PID do processo (quando registrado)

## Paths importantes
- Arquivos de configuração e dados:
  - `app_config.json`
  - `app_data.db`
- Diretório de downloads planejado:
  - `APP_DIR / downloads`
- Possível destino para instalação de arquivos:
  - `share/extensions`
  - `share/roms`
- Padrão de socket usado pela aplicação:
  - `%TEMP%/openmsx-default/socket.{pid}`

## Troubleshooting
- `openmsx.exe` não encontrado: verifique `Configuração` e se o `openmsx.exe` está no diretório configurado.
- `psutil` ausente: a aplicação funciona sem ele, mas a detecção precisa de PID pode ser menos confiável.
- Socket não encontrado: verifique se o `openMSX` iniciou corretamente e se o PID exibido é o correto.
- Problemas de rede ao baixar: verifique conectividade, certificados TLS e limites da API do provedor.

## Testes e CI
- Adicione testes unitários conforme o código evolui.
- Sugere-se configurar CI para linting e execução de testes (ex.: GitHub Actions).

## Contribuição
Contribuições são bem‑vindas. Abra issues para bugs/ideias e pull requests com mudanças bem descritas. Documente alterações que impactem o formato de `app_config.json` ou do banco `app_data.db`.

## Licença
Licença a ser definida pelo autor do repositório. Adicione um arquivo `LICENSE` quando decidir o tipo de licença.
