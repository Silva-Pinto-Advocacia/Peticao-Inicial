# Silva Pinto Advocacia — Gerador de Petições
**Sistema automatizado de petições iniciais para concursos públicos**

---

## Requisitos

- Python 3.10 ou superior
- LibreOffice (para conversão PDF — opcional mas recomendado)
- Chave de API Anthropic (console.anthropic.com)

---

## Instalação Local (Windows / Mac / Linux)

```bash
# 1. Instale as dependências
pip install -r requirements.txt

# 2. Execute o servidor
python app.py
```

Acesse: **http://localhost:5000**

---

## Deploy em Servidor (Produção)

### Opção 1 — Render.com (gratuito)

1. Crie conta em https://render.com
2. Clique em **New → Web Service**
3. Conecte seu repositório GitHub (faça upload deste projeto lá primeiro)
4. Configure:
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
5. Clique em **Deploy**
6. Você receberá um link público como `https://seu-app.onrender.com`

```bash
# Adicionar gunicorn ao requirements.txt para produção:
pip install gunicorn
```

### Opção 2 — Railway.app

1. Acesse https://railway.app
2. Novo projeto → Deploy from GitHub
3. Variável de ambiente: não necessária (a chave é inserida na interface)
4. Link público gerado automaticamente

### Opção 3 — VPS (DigitalOcean / Linode / Hetzner)

```bash
# Instalar dependências do sistema
sudo apt update && sudo apt install -y python3-pip libreoffice

# Instalar app
pip install -r requirements.txt gunicorn

# Rodar com gunicorn (porta 8000)
gunicorn -w 2 -b 0.0.0.0:8000 app:app

# Usar nginx como proxy reverso (recomendado)
```

---

## Estrutura do Projeto

```
peticoes_app/
├── app.py                    # Backend Flask principal
├── requirements.txt          # Dependências Python
├── README.md
├── templates/
│   └── index.html            # Interface web
├── scripts/
│   └── office/
│       ├── unpack.py         # Desempacota DOCX
│       ├── pack.py           # Empacota DOCX
│       └── soffice.py        # Conversão LibreOffice
├── uploads/                  # Arquivos temporários (criado automaticamente)
└── output/                   # ZIPs gerados (criado automaticamente)
```

---

## Uso

1. Acesse o link da aplicação
2. Insira sua **chave de API Anthropic** (sk-ant-api03-...)
3. Faça upload do **ZIP do cliente** contendo:
   - `modelo.docx` (template da petição)
   - Ficha do candidato (XLSX)
   - Documentos de identidade (PDF/JPG)
   - Relatórios técnicos das questões (PDF/DOCX)
4. Preencha os dados do caso
5. Clique em **Gerar Petição**
6. Baixe o ZIP final com petição e documentos renomeados

---

## Segurança

- A chave de API **não é armazenada** no servidor
- Os arquivos de upload são temporários (sessão única)
- Para produção, recomenda-se adicionar autenticação básica

---

## Suporte

OAB/RJ nº 189.781 — Dr. Casil da Silva Pinto
