import io

import openpyxl
from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from .classificador import classificar, reclassificar
from .models import Conta, Lancamento, Regra
from . import planilha


def planilha_exemplo(senha=None):
    """Monta um .xlsx no mesmo formato da planilha Gestão Financeira."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    r = wb.create_sheet("Regras")
    r.append(["Regras"]); r.append(["explicação"]); r.append([])
    r.append(["Palavra-chave", "Tipo de busca", "Seção", "Plano de Contas", "Observação", None, None, "Contas válidas"])
    contas = ["-> Compras de Aço", "-> Energia Elétrica", "-> Fretes", "-> Recebimento PIX", "-> Transferência entre contas"]
    regras = [
        ("PIX", "Contém", "Receita", "-> Recebimento PIX", ""),
        ("AÇO", "Contém", "Despesa", "-> Compras de Aço", ""),
        ("ACO", "Contém", "Despesa", "-> Compras de Aço", ""),
        ("ENEL", "Começa com", "Ambas", "-> Energia Elétrica", ""),
        ("FRETE", "Contém", "Despesa", "-> Fretes", ""),
        ("FRETE AÇO", "Contém", "Despesa", "-> Fretes", "exceção: frete de aço é frete"),
    ]
    for i in range(max(len(contas), len(regras))):
        linha = list(regras[i]) if i < len(regras) else [None] * 5
        linha += [None, None, contas[i] if i < len(contas) else None]
        r.append(linha)
    a = wb.create_sheet("2024R")
    a.append(["3.1", "-> Compras de Aço"])
    jan = wb.create_sheet("Jan")
    jan.append(["Janeiro"]); jan.append([]); jan.append([None, "Descrição", "Plano de Contas", "Valor"])
    jan.append([None, "PIX CLIENTE JOAO", None, 1500])           # receita, regra PIX
    jan.append([None, "ENEL CONTA LUZ", None, "200,50"])         # receita? ENEL Ambas
    jan.append([None, "DESPESAS"])
    jan.append([None, "AÇO GERDAU NF 123", "-> Compras de Aço", 9000])  # bate com a regra -> auto
    jan.append([None, "FRETE AÇO GERDAU", None, 300])            # exceção vence (ordem maior)
    jan.append([None, "PIX FORNECEDOR", None, 50])               # PIX é só receita -> sem regra
    jan.append([None, "ACOUGUE DO ZE", "-> Energia Elétrica", 80])  # digitado à mão -> manual
    jan.append([None, "EMPEDRADO", None, 10])
    fev = wb.create_sheet("Fev")
    fev.append([]); fev.append([]); fev.append([])
    fev.append([None, "DESPESAS"])
    fev.append([None, "enel fevereiro", None, 210])
    buf = io.BytesIO()
    wb.save(buf)
    if senha:
        import msoffcrypto.format.ooxml
        buf.seek(0)
        saida = io.BytesIO()
        msoffcrypto.format.ooxml.OOXMLFile(buf).encrypt(senha, saida)
        buf = saida
    buf.seek(0)
    return buf


class ClassificadorTests(TestCase):
    def setUp(self):
        self.aco = Conta.objects.create(nome="Aço")
        self.frete = Conta.objects.create(nome="Frete")
        Regra.objects.create(palavra="AÇO", conta=self.aco, secao="D", ordem=1)
        Regra.objects.create(palavra="FRETE", tipo="I", conta=self.frete, secao="A", ordem=2)

    def test_inicio_de_palavra(self):
        self.assertEqual(classificar("compra de aço nf 1", "D").conta, self.aco)
        self.assertEqual(classificar("NF-AÇOS LTDA", "D").conta, self.aco)
        self.assertIsNone(classificar("PEDAÇO", "D"))

    def test_secao(self):
        self.assertIsNone(classificar("AÇO", "R"))

    def test_comeca_com_e_prioridade(self):
        self.assertIsNone(classificar("PAGTO FRETE", "D"))  # FRETE é "começa com"
        self.assertEqual(classificar("FRETE AÇO", "D").conta, self.frete)  # ordem maior vence

    def test_acento_importa(self):
        self.assertIsNone(classificar("ACO", "D"))


class ImportacaoTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user("pai", password="senha-forte-123")

    def importar(self, **kw):
        wb = planilha.abrir_planilha(planilha_exemplo(), "")
        return planilha.importar(wb, 2026, usuario=self.user, coluna_valor="D", **kw)

    def test_importa_tudo(self):
        r = self.importar()
        self.assertEqual(r["regras"], 6)
        self.assertEqual(r["lancamentos"], 8)
        self.assertEqual(Conta.objects.get(nome="-> Compras de Aço").codigo, "3.1")
        self.assertTrue(Conta.objects.get(nome="-> Transferência entre contas").nao_entra_no_dre)

        get = lambda d: Lancamento.objects.get(descricao=d)
        self.assertEqual(get("PIX CLIENTE JOAO").secao, "R")
        self.assertEqual(get("PIX CLIENTE JOAO").conta.nome, "-> Recebimento PIX")
        self.assertEqual(str(get("ENEL CONTA LUZ").valor), "200.50")
        self.assertEqual(get("AÇO GERDAU NF 123").classificacao, "A")
        self.assertEqual(get("FRETE AÇO GERDAU").conta.nome, "-> Fretes")
        self.assertEqual(get("PIX FORNECEDOR").classificacao, "N")
        self.assertEqual(get("ACOUGUE DO ZE").classificacao, "M")
        self.assertEqual(get("enel fevereiro").mes, 2)
        self.assertEqual(get("enel fevereiro").secao, "D")
        # histórico registrado com usuário e motivo
        h = Lancamento.history.first()
        self.assertEqual(h.history_user, self.user)
        self.assertEqual(h.history_change_reason, "Importação da planilha")

    def test_reimportar_substitui(self):
        self.importar()
        self.importar()
        self.assertEqual(Lancamento.objects.count(), 8)
        self.assertEqual(Regra.objects.count(), 6)
        self.assertEqual(Lancamento.history.filter(history_type="-").count(), 8)

    def test_planilha_com_senha(self):
        try:
            arq = planilha_exemplo(senha="lmfinanceiro")
        except Exception as exc:  # versão do msoffcrypto sem suporte a criptografar
            self.skipTest(f"sem suporte para gerar planilha com senha: {exc}")
        with self.assertRaises(planilha.ErroImportacao):
            planilha.abrir_planilha(planilha_exemplo(senha="lmfinanceiro"), "errada")
        wb = planilha.abrir_planilha(arq, "lmfinanceiro")
        self.assertIn("Regras", wb.sheetnames)

    def test_exportar(self):
        self.importar()
        wb = openpyxl.load_workbook(planilha.exportar(2026))
        self.assertEqual(wb.sheetnames, ["Lançamentos 2026", "Regras"])
        self.assertEqual(wb["Lançamentos 2026"].max_row, 9)


class TelasTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user("vilma", password="senha-forte-123")
        wb = planilha.abrir_planilha(planilha_exemplo(), "")
        planilha.importar(wb, 2026, usuario=self.user, coluna_valor="D")

    def test_exige_login(self):
        for nome in ["lancamentos", "regras", "historico", "resumo", "importar", "contas"]:
            resp = self.client.get(reverse(nome))
            self.assertEqual(resp.status_code, 302, nome)
            self.assertIn("/entrar/", resp["Location"])

    def test_paginas_abrem(self):
        self.client.force_login(self.user)
        lanc = Lancamento.objects.first()
        regra = Regra.objects.first()
        urls = [reverse(n) for n in ["lancamentos", "regras", "historico", "resumo", "importar", "contas",
                                     "lancamento_novo", "regra_nova", "conta_nova"]]
        urls += [reverse("lancamento_editar", args=[lanc.pk]), reverse("regra_editar", args=[regra.pk]),
                 reverse("regras") + "?testar=FRETE+AÇO&secao_teste=D",
                 reverse("sugerir") + "?descricao=frete&secao=D",
                 reverse("historico") + "?modelo=lancamento&tipo=%2B&q=PIX",
                 reverse("lancamentos") + "?classificacao=N&mes=1"]
        for url in urls:
            self.assertEqual(self.client.get(url).status_code, 200, url)
        self.assertEqual(self.client.get(reverse("exportar")).status_code, 200)

    def test_correcao_manual_e_historico(self):
        self.client.force_login(self.user)
        lanc = Lancamento.objects.get(descricao="PIX FORNECEDOR")
        energia = Conta.objects.get(nome="-> Energia Elétrica")
        resp = self.client.post(reverse("lancamento_conta", args=[lanc.pk]), {"conta": energia.pk})
        self.assertContains(resp, f'id="l{lanc.pk}"')
        lanc.refresh_from_db()
        self.assertEqual((lanc.conta, lanc.classificacao), (energia, "M"))
        h = lanc.history.first()
        self.assertEqual(h.history_user, self.user)
        self.assertEqual(h.history_change_reason, "Correção manual da conta")
        # aparece no histórico com a diferença
        resp = self.client.get(reverse("historico") + "?q=Correção")
        self.assertContains(resp, "Energia Elétrica")
        self.assertContains(resp, "vilma")
        # voltar para automática
        self.client.post(reverse("lancamento_conta", args=[lanc.pk]), {"conta": "auto"})
        lanc.refresh_from_db()
        self.assertEqual(lanc.classificacao, "N")

    def test_nova_regra_reclassifica(self):
        self.client.force_login(self.user)
        conta = Conta.objects.get(nome="-> Fretes")
        resp = self.client.post(reverse("regra_nova"), {
            "palavra": "pix", "tipo": "C", "secao": "D", "conta": conta.pk, "ordem": 100, "ativa": "on"})
        self.assertEqual(resp.status_code, 302)
        lanc = Lancamento.objects.get(descricao="PIX FORNECEDOR")
        self.assertEqual((lanc.conta, lanc.classificacao), (conta, "A"))
        # manual não é tocado
        self.assertEqual(Lancamento.objects.get(descricao="ACOUGUE DO ZE").classificacao, "M")
        self.assertEqual(lanc.history.first().history_user, self.user)

    def test_novo_lancamento(self):
        self.client.force_login(self.user)
        resp = self.client.post(reverse("lancamento_novo"), {
            "ano": 2026, "mes": 3, "secao": "D", "descricao": "Frete para obra", "valor": "120.00", "conta": ""})
        self.assertEqual(resp.status_code, 302)
        l = Lancamento.objects.get(descricao="Frete para obra")
        self.assertEqual((l.conta.nome, l.classificacao), ("-> Fretes", "A"))
        self.assertEqual(l.history.first().history_user, self.user)

    def test_login(self):
        resp = self.client.post(reverse("login"), {"username": "vilma", "password": "senha-forte-123"})
        self.assertEqual(resp.status_code, 302)
        self.client.logout()
        resp = self.client.post(reverse("login"), {"username": "vilma", "password": "errada"})
        self.assertContains(resp, "incorretos")

    def test_importar_pela_tela(self):
        self.client.force_login(self.user)
        arq = planilha_exemplo()
        arq.name = "Gestao.xlsx"
        resp = self.client.post(reverse("importar"), {
            "arquivo": arq, "ano": 2025, "importar_regras": "on", "importar_lancamentos": "on", "coluna_valor": "D"})
        self.assertContains(resp, "Lançamentos importados: <b>8</b>")
        self.assertEqual(Lancamento.objects.filter(ano=2025).count(), 8)
