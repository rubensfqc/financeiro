# Gestão Financeira – classificação automática

Webapp Django que faz o mesmo que a planilha *Gestão Financeira*: cada lançamento recebe a conta do plano de contas automaticamente, a partir de palavras-chave da descrição. Tem login com senha e guarda o histórico de todas as alterações (quem, quando, o que mudou).

## O que tem

- **Login obrigatório** em todas as telas, troca de senha, sessão de 8 h. Usuários são criados no admin (`/admin/`, menu *Usuários* para quem é staff).
- **Lançamentos**: filtro por ano, mês, seção, conta, classificação e texto. A conta pode ser trocada direto na tabela; a escolha manual fica travada e as regras não mexem mais nela (dá para voltar para automática). Linhas sem regra ficam em amarelo, com atalho “+ regra”.
- **Regras**: mesma lógica da aba Regras da planilha (veja abaixo), com um testador (“qual regra pega esta descrição?”), contagem de uso e lista das descrições sem regra mais frequentes. Salvar, editar ou excluir uma regra reclassifica na hora todos os lançamentos automáticos.
- **Contas**: plano de contas com código, marcação “fora do DRE” (transferências entre contas) e ativa/inativa.
- **Resumo**: totais por conta e por mês (ou quantidades, se os valores não forem importados).
- **Histórico**: toda criação, alteração e exclusão de lançamentos, regras e contas, com usuário, data, motivo e os campos alterados (antes → depois). Filtros por tipo, ação, usuário, período e texto. Cada lançamento e regra também mostra o próprio histórico.
- **Importar / exportar .xlsx**: importa a planilha atual (inclusive com senha de abertura) e exporta lançamentos + regras.

### Regras de classificação

- Maiúsculas/minúsculas não importam; acentos importam (cadastre AÇO **e** ACO).
- **Contém**: a palavra-chave aparece no início de qualquer palavra da descrição (PEDRA pega “PEDRA BRITADA”, não pega “EMPEDRADO”). Pontuação conta como separador: “NF-AÇO” também casa com AÇO.
- **Começa com**: a descrição começa com a palavra-chave.
- **Seção**: Receita, Despesa ou Ambas.
- Se várias regras servirem, vence a de **maior ordem** (na planilha, a que estava mais abaixo). Exceções recebem número alto.

## Rodar localmente

```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python manage.py migrate
python manage.py createsuperuser
python manage.py runserver
```

Abra http://127.0.0.1:8000, entre e vá em **Importar**.

### Importar a planilha atual

Pela tela **Importar** (envia o arquivo, a senha e o ano) ou pela linha de comando:

```bash
python manage.py importar_planilha Gestao_Financeira_2026.xlsx --ano 2026 --usuario admin
# pede a senha se a planilha tiver; --coluna-valor D para importar também os valores
```

O que é lido:

- **Aba Regras**: regras a partir da linha 5 (A palavra, B tipo, C seção, D conta, E observação) e contas válidas na coluna H.
- **Aba 2024R** (opcional): código da conta (coluna A) para cada nome (coluna B).
- **Abas Jan…Dez**: descrição na coluna B e conta na coluna C, a partir da linha 4. Receitas ficam acima da linha “DESPESAS” e despesas abaixo dela.
- **Valores**: informe a letra da coluna em *Coluna do valor*. Em branco, os valores não são importados e o Resumo mostra quantidades.
- **Correções manuais**: quando a conta da coluna C é diferente do que as regras dariam (alguém digitou por cima), o lançamento entra como **Manual**.
- **Reimportar**: importar de novo o mesmo ano substitui os lançamentos desse ano e as regras. O que foi apagado continua no histórico.

Outros comandos: `python manage.py reclassificar` e `python manage.py test financas` (15 testes).

## Publicar no PythonAnywhere

1. No console Bash:
   ```bash
   git clone <seu-repo> ~/financeiro      # ou envie o zip e descompacte
   cd ~/financeiro
   mkvirtualenv financeiro --python=python3.12
   pip install -r requirements.txt
   cp .env.example .env && nano .env      # SECRET_KEY, DEBUG=0, ALLOWED_HOSTS
   python manage.py migrate
   python manage.py collectstatic --noinput
   python manage.py createsuperuser
   ```
2. Aba **Web** → *Add a new web app* → *Manual configuration* (Python 3.12).
   - **Virtualenv**: `/home/<usuario>/.virtualenvs/financeiro`
   - **Static files**: URL `/static/` → `/home/<usuario>/financeiro/staticfiles`
   - **Force HTTPS**: ligado
3. Edite o **WSGI configuration file**:
   ```python
   import os, sys
   path = "/home/<usuario>/financeiro"
   if path not in sys.path:
       sys.path.insert(0, path)
   os.environ["DJANGO_SETTINGS_MODULE"] = "config.settings"
   from django.core.wsgi import get_wsgi_application
   application = get_wsgi_application()
   ```
4. **Reload**. O `.env` é lido automaticamente pelo `settings.py`.

**Backup**: o banco é o arquivo `db.sqlite3`. Copie-o periodicamente, por exemplo com uma *Scheduled task* diária `cp ~/financeiro/db.sqlite3 ~/backups/db-$(date +%F).sqlite3`. Para mais usuários simultâneos, troque `DATABASES` para o MySQL do PythonAnywhere.

## Estrutura

```
config/                 settings, urls, wsgi
financas/
  models.py             Conta, Regra, Lancamento (+ histórico via django-simple-history)
  classificador.py      lógica das regras e reclassificação
  planilha.py           importação/exportação .xlsx
  views.py, urls.py, forms.py, admin.py
  templates/, static/   telas (Django templates + HTMX, CSS próprio)
  management/commands/  importar_planilha, reclassificar
  tests.py
```
