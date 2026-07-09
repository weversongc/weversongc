# 🤖 Weversongc - Dashboard de Automações

Plataforma web centralizadora de automações financeiras. Roda 100% na nuvem (gratuito) com Streamlit + Supabase.

🔗 Acesse: **https://weversongc.streamlit.app**

## 🛠️ Ferramentas

| Aba | Função |
|-----|--------|
| 📄 Conversor PDF → Excel | Transforma PDFs em abas estruturadas no Excel |
| 📊 Conciliação Bancária | Cruza Braslog (XLSX) com extrato Sicoob (CSV) |
| 💸 Diferenças Braslog vs MH | Analisa transferências entre as empresas |

## 🚀 Rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

Configure as credenciais copiando `.streamlit/secrets.toml.example` para `.streamlit/secrets.toml` e preenchendo sua chave anon do Supabase.

## ☁️ Deploy (Streamlit Community Cloud)

1. Suba este repositório para o GitHub
2. Acesse https://sharing.streamlit.io → New app
3. Selecione o repositório e o arquivo `app.py`
4. Em **Settings → Secrets**, cole:
   ```toml
   SUPABASE_URL = "https://iwwttnusktosuaaiusoh.supabase.co"
   SUPABASE_KEY = "sua_chave_anon"
   NOME_BUCKET = "arquivo-usuários"
   ```
5. Deploy → o app fica disponível em `weversongc.streamlit.app`

## 📂 Estrutura

```
app.py                  -> aplicação principal (roteamento entre ferramentas)
requirements.txt        -> dependências
.streamlit/             -> configuração e secrets
referencia/             -> scripts locais originais (referência)
```

## 🔒 Segurança

As credenciais do Supabase **nunca** ficam no código — são lidas via `st.secrets`. O arquivo `secrets.toml` está no `.gitignore`.
