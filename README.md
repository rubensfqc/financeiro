# Gestão Financeira – classificação automática

Webapp Django que faz o mesmo que a planilha *Gestão Financeira*: cada lançamento recebe a conta do plano de contas automaticamente, a partir de palavras-chave da descrição. Tem login com senha e guarda o histórico de todas as alterações (quem, quando, o que mudou).

## O que tem

- **Login obrigatório** em todas as telas, troca de senha, sessão de 8 h. Usuários são criados no admin (`/admin/`, menu *Usuários* para quem é staff).
- **Lançamentos**: os filtros ficam guardados (por usuário) ao trocar de aba e voltar; “Limpar” zera. Filtro por ano, mês, seção, conta, classificação e texto. A conta pode ser trocada direto na tabela; a escolha manual fica travada e as regras não mexem mais nela (dá para voltar para automática). Linhas sem regra ficam em amarelo, com atalho “+ regra”.
- **Regras**: mesma lógica da aba Regras da planilha (veja abaixo), com um testador (“qual regra pega esta descrição?”), contagem de uso e lista das descrições sem regra mais frequentes. Salvar, editar ou excluir uma regra reclassifica na hora todos os lançamentos automáticos.
- **Bancos**: cadastro de bancos/caixas (nome, coluna da planilha, instituição, agência, número da conta, titular), com totais do ano e histórico. Cada lançamento mostra o banco.
- **Contas**: plano de contas com código, marcação “fora do DRE” (transferências entre contas) e ativa/inativa.
- **Dashboard**: visão consolidada por trimestre (T1–T4) ou ano — veja a seção abaixo.
- **Resumo**: totais por conta e por mês (ou quantidades, se os valores não forem importados).
- **Histórico**: toda criação, alteração e exclusão de lançamentos, regras e contas, com usuário, data, motivo e os campos alterados (antes → depois). Filtros por tipo, ação, usuário, período e texto. Cada lançamento e regra também mostra o próprio histórico.
- **Importar / exportar .xlsx**: importa a planilha atual (inclusive com senha de abertura) e exporta lançamentos + regras.

### Regras de classificação

- Maiúsculas/minúsculas não importam; acentos importam (cadastre AÇO **e** ACO).
- **Contém**: a palavra-chave aparece no início de qualquer palavra da descrição (PEDRA pega “PEDRA BRITADA”, não pega “EMPEDRADO”). Pontuação conta como separador: “NF-AÇO” também casa com AÇO.
- **Começa com**: a descrição começa com a palavra-chave.
- **Seção**: Receita, Despesa ou Ambas.
- Se várias regras servirem, vence a de **maior ordem** (na planilha, a que estava mais abaixo). Exceções recebem número alto.

## Dashboard

Página inicial após o login. Escolha o ano e o trimestre (ou o ano inteiro). Usa o plano de contas da empresa (grupos 1 a 7, já cadastrado pela migração; `python manage.py carregar_plano_contas` acerta os códigos se alguém mudar).

- **Cartões**: receita (grupo 1), custos variáveis (3), margem de contribuição em R$ e %, custos fixos (5), resultado operacional em R$ e %, ponto de equilíbrio. Cada um comparado com o trimestre anterior e com o mesmo trimestre do ano anterior.
- **DRE do período**: inclui receitas (2) e despesas (7) não operacionais e o saldo final.
- **Tendência**: receita e custos dos últimos 8 trimestres, margens em %, e o mês a mês do ano com a receita do ano anterior.
- **Pré-moldados**: compras de aço, cimento e mercadorias diversas por trimestre, e o peso na receita de aço, cimento, mercadorias, logística (combustível + fretes), mão de obra (MO avulsa + pessoal 5.1–5.8) e impostos.
- **Detalhamento**: 10 maiores despesas com variação, receita por forma de recebimento.
- **Alertas automáticos**: prejuízo, margem caindo 3 p.p. ou mais, indicador que muda 2 p.p. ou mais, conta que sobe 30% ou mais, receita 15% acima/abaixo do ano anterior.
- **Qualidade dos dados**: lançamentos sem conta, sem valor, ou em contas sem código.

Os valores são somados em módulo (o grupo da conta define se é receita ou custo). Contas “fora do DRE” (transferências) não entram.

**Ver com dados de exemplo**: `python manage.py dados_demo` cria lançamentos fictícios (marcados “[DEMO]”) para o ano passado e o atual; `python manage.py dados_demo --apagar` remove. Ele se recusa a rodar se já houver lançamentos reais (use `--forcar` para misturar).

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
# pede a senha se a planilha tiver; valor da coluna D (troque com --coluna-valor)
# --apenas-valores: só preenche os valores dos lançamentos já importados
```

O que é lido:

- **Aba Regras**: regras a partir da linha 5 (A palavra, B tipo, C seção, D conta, E observação) e contas válidas na coluna H.
- **Abas Jan…Dez**: descrição na coluna B e conta na coluna C, a partir da linha 4. Receitas ficam acima da linha “DESPESAS” e despesas abaixo dela.
- **Valores**: coluna D por padrão (dá para trocar a letra na tela). Aceita número, texto como “R$ 1.500,00” e fórmulas simples como `=600+150`, mesmo quando o arquivo não guardou o resultado calculado.
- **Plano de contas = coluna C**: o nome da coluna C é casado com o plano de contas da empresa ignorando seta (`->`), código na frente, espaços, maiúsculas e acentos. “->Compras de Aço”, “compras de aco” e “3.1 -> Compras de Aço” viram a mesma conta 3.1. Nomes que não existem no plano entram sem código, ficam fora do dashboard e aparecem num aviso em **Plano de contas**, onde dá para mesclá-los com a conta certa.
- **Banco**: cada banco tem uma coluna da planilha (inicialmente E SICOOB LM, F SICOOB RM, G ITAÚ, H BRADESCO, I SICOOB RF, J LM SIST CX ECON, K RM). O lançamento fica com o banco cuja coluna tem valor diferente de zero; sem valor em nenhuma, fica sem banco. Se a coluna D estiver vazia, o valor vem da coluna do banco. Bancos novos, coluna, agência e número da conta são editados em **Bancos**.
- **Fórmulas**: quando o arquivo não guardou o resultado das fórmulas (ex.: `=SOMA(E5:K5)` na coluna D), o app calcula a partir das células (referências, intervalos, SOMA/SUM e + − × ÷).
- **Só atualizar valores**: a opção “Só atualizar valor e banco” preenche valor e banco dos lançamentos já importados (pelo mês, linha e descrição) sem mexer em plano de contas, regras ou correções feitas no app.
- **Correções manuais**: quando a conta da coluna C é diferente do que as regras dariam (alguém digitou por cima), o lançamento entra como **Manual**.
- **Reimportar**: importar de novo o mesmo ano substitui os lançamentos desse ano e as regras. O que foi apagado continua no histórico.

Outros comandos: `python manage.py reclassificar`, `python manage.py carregar_plano_contas` e `python manage.py test financas` (35 testes).

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
  plano_contas.py       plano de contas (grupos 1–7) e agrupamentos dos indicadores
  contas.py             casa os nomes da coluna "Plano de Contas" com o plano; unifica/mescla contas
  dashboard.py          cálculos do dashboard
  planilha.py           importação/exportação .xlsx
  views.py, urls.py, forms.py, admin.py
  templates/, static/   telas (Django templates + HTMX, CSS próprio)
  management/commands/  importar_planilha, reclassificar, carregar_plano_contas, dados_demo
  tests.py
```
