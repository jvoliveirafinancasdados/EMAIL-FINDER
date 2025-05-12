Ferramenta GUI para analisar emails e contar palavras-chave, gerando planilha Excel com resultados.

## Funcionalidades
- Conexão segura com servidores IMAP (Outlook/Office365)
- Extração e processamento de emails
- Contagem de palavras-chave no corpo das mensagens
- Geração automática de planilha Excel
- Interface amigável com feedback visual
- Tratamento de erros detalhado

## Requisitos
- Python 3.8+
- Conta de email Outlook/Office365

## Instalação
1. Clone o repositório:
```bash
git clone https://github.com/seu-usuario/email-keyword-analyzer.git
cd email-keyword-analyzer
```

2. Instale as dependências:
```bash
pip install flet pandas
```

3. Execute o aplicativo:
```bash
python main.py
```

## Uso
1. Insira suas credenciais de email
2. Digite palavras-chave separadas por vírgula
3. Clique em "Analisar Emails"
4. Aguarde o processamento
5. Um arquivo `resultados_emails.xlsx` será gerado

## Notas de Segurança
- Use uma **Senha de Aplicativo** ao invés da senha principal
- O código não armazena credenciais
- Recomendado para contas de teste

## Limitações
- Processa no máximo 50 emails por execução
- Suporte apenas para Outlook/Office365
- Requer conexão com internet

## Erros Comuns
- **Login falhou**: Verifique as credenciais e habilite autenticação de dois fatores
- **Sem conexão**: Verifique sua conexão com a internet
- **Arquivo bloqueado**: Feche o Excel antes de executar

## Licença
MIT License - veja [LICENSE](LICENSE) para detalhes

