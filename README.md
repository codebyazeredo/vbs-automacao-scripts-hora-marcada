# Agendamento por Horários Fixo de BATs

Gerenciador leve para execução de scripts `.BAT` no Windows em horários específicos do dia, com **suporte a execução paralela**, controle de concorrência por arquivos de trava (`.lock`) e retenção automática de logs.

Diferente do modelo baseado em intervalos de minutos, este script dispara tarefas paralelas assim que seus horários programados são atingidos, monitorando o término de cada processo através de um mecanismo de *polling*.

---

## Estrutura do Projeto

```text
vbs/
├── bat/                 # Pasta para armazenar os arquivos BAT
│   └── teste.bat
├── logs/                # Histórico de execuções normais e erros (automático)
│   ├── execucao_2026-09-03.log
│   └── erros_2026-09-03.log
├── locks/               # Travas temporárias para impedir execuções duplicadas
├── tmp/                 # Scripts wrappers (.cmd) e retornos temporários
├── .gitignore
├── agendador.vbs        # Script principal gerenciador
├── config_horario.ini   # Mapeamento de scripts e lista de horários (HH:MM)
└── estado_horario.ini   # Registro de tarefas concluídas no dia (automático)

```

---

## Arquivos de Configuração

### 1. `config_horario.ini`

Define os arquivos `.BAT` a serem executados e uma lista de horários no formato 24 horas (`HH:MM`), separados por vírgula.

```ini
script_backup.bat=08:00, 12:00, 18:00
relatorio_diario.bat=08:30
limpeza_tmp.bat=23:50

```

* **Lado esquerdo (`=`):** Nome exato do arquivo `.BAT` presente na pasta `bat/`.
* **Lado direito (`=`):** Lista de horários fixos de execução separados por vírgula.

### 2. `estado_horario.ini`

Gerado e atualizado automaticamente pelo script. Armazena a data (`YYYYMMDD`) da última execução bem-sucedida ou com falha de cada chave `nome_do_script.bat|HH:MM`, evitando que a mesma tarefa rode mais de uma vez no mesmo dia.

---

## Pastas do Sistema

* **`bat/`**: Diretório onde ficam armazenados os scripts `.bat` executáveis.
* **`logs/`**: Armazena os registros de saída. Arquivos com mais de 60 dias são removidos automaticamente pelo script.
* `execucao_YYYY-MM-DD.log`: Contém o resumo de execuções bem-sucedidas.
* `erros_YYYY-MM-DD.log`: Registra falhas, parâmetros inválidos ou relatórios de erro retornados pelos scripts.


* **`locks/`**: Arquivos `.lock` temporários para que uma mesma tarefa não rode em duplicidade se a execução anterior ainda estiver em andamento.
* **`tmp/`**: Armazena os arquivos wrappers `.cmd` que executam as rotinas em segundo plano sem exibir janelas do prompt, capturando o código de retorno (`errorlevel`) e o texto de saída.

---

## Fluxo de Funcionamento

O script foi projetado para ser chamado periodicamente (ex: a cada 1 ou 2 minutos) via Agendador de Tarefas do Windows.

1. **Leitura e Validação:** O script lê o `config_horario.ini` e compara a hora atual com cada horário agendado.
2. **Disparo em Paralelo:** Se o horário atual for igual ou maior que o agendado e a tarefa ainda não tiver rodado hoje (verificado via `estado_horario.ini`), um script wrapper é gerado em `tmp/` e executado de forma assíncrona.
3. **Monitoramento (*Polling*):** A cada 2 segundos (`INTERVALO_POLL_MS = 2000`), o VBScript checa a pasta `tmp/` para verificar se os arquivos de conclusão foram gerados.
4. **Finalização:** Quando a rotina encerra, o código de erro e as mensagens de saída são capturados, os arquivos temporários e de trava são excluídos, o `estado_horario.ini` é atualizado e o log é gravado.

---

## Configuração no Agendador de Tarefas do Windows

1. Pressione `Win + R`, digite **`taskschd.msc`** e pressione **Enter**.
2. No menu lateral, clique em **Criar Tarefa**.
3. Configure as abas:

### Aba Geral

* **Nome:** `Executor por Horario - VBS`
* Marque: **Executar estando o usuário conectado ou não**
* Marque: **Executar com privilégios mais altos**

### Aba Disparadores

* Clique em **Novo...** e configure para iniciar **Diariamente**.
* Em *Configurações avançadas*, marque **Repetir a tarefa a cada:** `1 minuto` (ou 5 minutos).
* Na opção *por um período de:*, escolha **Indefinidamente**.

### Aba Ações

* Clique em **Nova...** e selecione **Iniciar um programa**.
* **Programa/script:** `wscript.exe`
* **Adicionar argumentos:** `"C:\CAMINHO\vbs\agendador.vbs"` *(substitua pelo seu caminho real)*
* **Iniciar em:** `C:\CAMINHO\vbs` *(substitua pelo caminho real sem aspas)*

### Aba Configurações

* Em *Se a tarefa já estiver em execução, a seguinte regra se aplica:*, selecione **Não iniciar uma nova instância**.

---

## Testando a Aplicação

1. Certifique-se de que o arquivo `config_horario.ini` possui um horário atual ou passado recente para teste.
2. Execute o arquivo VBScript via linha de comando ou duplo clique.
3. Verifique o diretório `logs/` para conferir a gravação das entradas e a pasta `estado_horario.ini` para confirmar a gravação do registro do dia.