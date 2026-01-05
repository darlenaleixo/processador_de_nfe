# 🚀 Processador de NFe - Versão Profissional

Uma aplicação de desktop completa, desenvolvida em Python com Tkinter, para automatizar o processo de backup, organização e análise de arquivos XML de NF-e baseado na **chave de acesso**.

## 📖 Descrição

Este programa foi criado para resolver um problema comum: a gestão manual e trabalhosa de arquivos XML de notas fiscais. A aplicação automatiza todo o fluxo de trabalho, desde a localização dos arquivos baseado na **chave de acesso da NFe** até o upload seguro para a nuvem e a notificação por e-mail.

### 🔥 Novidade: Filtragem por Chave de Acesso

A principal inovação desta versão é a **filtragem inteligente baseada na chave de acesso da NFe**:
- ✅ Analisa o mês/ano real de emissão (posições 3-6 da chave de 44 dígitos)
- ✅ Independe da data de modificação dos arquivos no computador
- ✅ Maior precisão na seleção das notas do período correto
- ✅ Funciona mesmo se arquivos foram movidos ou copiados

## 🏗️ Estrutura do Projeto

```
tentativa_de_app/
│
├── app.py                 # Arquivo principal (entrada do programa)
├── requirements.txt       # Dependências do projeto
├── README.md             # Esta documentação
│
├── gui/                  # Interface gráfica
│   ├── __init__.py
│   └── main_window.py    # Janela principal da aplicação
│
├── nfe/                  # Processamento de NFe
│   ├── __init__.py
│   └── parser.py         # Lógica de análise dos XMLs
│
├── services/             # Serviços auxiliares
│   ├── __init__.py
│   ├── email_service.py  # Envio de e-mails
│   ├── rclone_service.py # Upload para nuvem
│   └── scheduler_service.py # Agendamento Windows
│
└── config/               # Configurações
    ├── __init__.py
    └── settings.py       # Gerenciamento de configurações
```

## ✨ Funcionalidades

- **Interface Gráfica Completa:** Construída com Tkinter, organizada em abas
- **Filtragem Inteligente por Chave de Acesso:** Analisa o mês/ano real da emissão
- **Organização Automática:** Copia arquivos para pasta organizada por período
- **Extração de Dados:** Gera relatório CSV detalhado com informações das NFes
- **Compactação:** Cria arquivo ZIP dos XMLs do mês
- **Upload para Nuvem:** Usa rclone para envio ao Google Drive
- **Notificações por E-mail:** Relatório automático de sucesso/falha
- **Agendamento:** Execução automática mensal via Agendador do Windows
- **Arquitetura Modular:** Código organizado profissionalmente

## ⚙️ Instalação

1. **Clone/baixe o projeto:**
   ```bash
   git clone <url-do-repositorio>
   cd tentativa_de_app
   ```

2. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure o rclone** (se usar upload para nuvem):
   ```bash
   rclone config
   ```

4. **Execute a aplicação:**
   ```bash
   python app.py
   ```

## 🔧 Configuração

1. **Abra a aplicação** e vá para a aba "Configurações"
2. **Configure os caminhos:**
   - Pasta Origem: onde estão os XMLs da NFe
   - Pasta Destino: onde salvar os backups
   - Caminho Rclone: localização do executável rclone.exe

3. **Configure o e-mail:**
   - Servidor SMTP, porta, credenciais
   - E-mails de origem e destino

4. **Salve as configurações** clicando em "Salvar Configurações"

## 🚀 Como Usar

### Execução Manual
1. Vá para a aba "Execução"
2. Clique em "Executar Backup Agora"
3. Acompanhe o progresso nos logs

### Agendamento Automático
1. Vá para a aba "Agendamento"
2. Clique em "Criar Agendamento"
3. O backup será executado automaticamente todo dia 1º do mês

## 🔍 Como Funciona a Nova Filtragem

A chave de acesso da NFe possui 44 dígitos organizados assim:

| Posições | Conteúdo | Exemplo |
|----------|----------|---------|
| 1-2 | Código UF | 35 (SP) |
| **3-6** | **AAMM (Ano/Mês)** | **2407** (Jul/2024) |
| 7-20 | CNPJ do emitente | 12345678000155 |
| 21-44 | Outros dados | ... |

O programa:
1. Lê **todos** os arquivos XML da pasta
2. Extrai a chave de acesso de cada um
3. Verifica se as posições 3-6 correspondem ao mês anterior
4. Copia apenas as NFes do período correto

## 📁 Arquivos Gerados

Para cada execução, são criados:
- **Pasta organizada:** `2024-07_JULHO/` com os XMLs do mês
- **Arquivo ZIP:** `NFEs_JUL_2024.zip` com todos os XMLs
- **Relatório CSV:** `Resumo_Detalhado_NFEs_2024-07_JULHO.csv`
- **Log de execução:** `log_copia_nfe.log`

## 🔄 Versionamento

- **v1.0:** Versão original (filtrava por data de modificação)
- **v2.0:** Nova versão com filtragem por chave de acesso + arquitetura modular

## 🤝 Contribuição

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 📄 Licença

Este projeto é de uso interno e proprietário.

## 📞 Suporte

Para dúvidas ou problemas:
- E-mail: solracinformatica@gmail.com
- Verifique os logs da aplicação em caso de erro