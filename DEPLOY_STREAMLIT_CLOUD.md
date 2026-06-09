# Publicar o QC-INI no Streamlit Cloud — modo PROTEGIDO (grátis)

> Objetivo: link público para testar fora da sua rede, **sem expor o código**
> e **restringindo quem entra**. Custo zero.

## Como a proteção funciona (3 camadas)

| Camada | O que protege | Como |
|---|---|---|
| 1. **Repositório PRIVADO no GitHub** | Ninguém vê/copia seu código | Deploy a partir de repo privado (Streamlit Cloud suporta no plano grátis) |
| 2. **Login do próprio app** | Só entra quem tem usuário/senha | Já existe (tela de login). A dica de senha foi escondida na nuvem |
| 3. **App restrito a e-mails (opcional)** | Só e-mails autorizados abrem a URL | Configuração "Specific viewers" no Streamlit Cloud |

> Importante: o **código nunca é exposto pelo app em si** — o usuário só vê a
> interface, nunca o Python. O único lugar onde o código apareceria é o link do
> GitHub no menu — e por isso (a) o repositório é **privado** e (b) o menu foi
> escondido (`toolbarMode = "minimal"`).

---

## ⚠️ Limitação do plano grátis (honestidade)
O banco SQLite é **efêmero** no Streamlit Cloud: reinicia quando o app hiberna.
Para **testar**, é perfeito (volta aos dados de demonstração). Para **produção**
com clientes pagantes, é preciso banco persistente (PostgreSQL) — fase seguinte.

---

## Passo a passo

### 1. Conta no GitHub
- Crie em https://github.com (se ainda não tiver).

### 2. Criar um repositório PRIVADO
- No GitHub: **New repository** → nome `qc-ini` → marque **Private** → Create.

### 3. Enviar o projeto (terminal, dentro da pasta QC_INI)
```bash
git init
git add .
git commit -m "QC-INI - sistema de controle de qualidade"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/qc-ini.git
git push -u origin main
```
> O `.gitignore` já evita subir `.venv`, banco local e cache.

### 4. Deploy no Streamlit Cloud
- Acesse https://share.streamlit.io e entre com o GitHub.
- Autorize o acesso aos **repositórios privados** quando pedir.
- **New app** → selecione:
  - **Repository:** `SEU_USUARIO/qc-ini` (privado)
  - **Branch:** `main`
  - **Main file path:** `qclab/app.py`
- **Deploy**. Em 1–2 min sai o link, ex.: `https://qc-ini.streamlit.app`.

### 5. Restringir quem acessa (recomendado)
No painel do app no Streamlit Cloud:
- **Settings → Sharing** → mude de "Public" para **"Specific viewers"**.
- Adicione os **e-mails** (Google) das pessoas que podem entrar.
- Só esses e-mails conseguem abrir a URL — os demais recebem "acesso negado".

### 6. Trocar a senha padrão (importante!)
Antes de compartilhar, troque a senha de `TESTE05` por uma forte:
- Pela tela **Configurações** do app, ou
- Recriando o usuário em `seed_data.py` e rodando `python -m database` localmente,
  depois `git push`.

### 7. Usar como "app" no celular
- Abra o link no Chrome do celular → menu **⋮** → **Adicionar à tela inicial**.

---

## Resumo da segurança
- ✅ Código **privado** (repo privado + menu escondido).
- ✅ Acesso por **login** do app (senha forte).
- ✅ (Opcional) Acesso por **e-mails autorizados**.
- ⚠️ Dados **não persistem** no grátis — só para teste.

Para produção real (multi-lab, dados persistentes, LGPD), seguimos para
PostgreSQL + hospedagem dedicada quando você decidir avançar.
