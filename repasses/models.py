"""
Modelos de dados do sistema de repasses médicos.

Por enquanto temos o cadastro de Médicos, que é a base de qualquer repasse.
Os demais modelos (procedimentos, regras de honorário, repasses gerados,
integrações com MedPlus e OMIE) serão adicionados conforme recebermos os
modelos de relatório e de importação.
"""

import re

from django.db import models

_RE_SUFIXO = re.compile(r'\s*\([^)]*\)\s*$')


class Medico(models.Model):
    """Cadastro central de médicos — a 'fonte única da verdade' dos profissionais."""

    CATEGORIA_RESIDENTE = 'residente'
    CATEGORIA_FELLOW = 'fellow'
    CATEGORIA_PRECEPTOR = 'preceptor'
    CATEGORIA_ANESTESISTA = 'anestesista'
    CATEGORIA_MEDICO = 'medico'
    CATEGORIA_CHOICES = [
        (CATEGORIA_MEDICO, 'Médico / Outros'),
        (CATEGORIA_FELLOW, 'Fellow'),
        (CATEGORIA_RESIDENTE, 'Residente'),
        (CATEGORIA_PRECEPTOR, 'Preceptor'),
        (CATEGORIA_ANESTESISTA, 'Anestesista'),
    ]

    nome = models.CharField('Nome completo', max_length=200)
    categoria = models.CharField('Categoria', max_length=20, choices=CATEGORIA_CHOICES,
                                 default=CATEGORIA_MEDICO)
    razao_social = models.CharField('Razão Social', max_length=200, blank=True,
                                    help_text='Usada como Fornecedor/Cliente na OMIE.')
    regra_obs = models.CharField('Regra / observação da planilha', max_length=255, blank=True,
                                 help_text='Ex.: "R$ 800,00 /semana", "Não recebem em catarata".')

    # Papéis (um médico pode acumular): alimentam as listas de seleção das cirurgias
    eh_fellow = models.BooleanField('É fellow?', default=False)
    eh_preceptor = models.BooleanField('É preceptor?', default=False)
    eh_anestesista = models.BooleanField('É anestesista?', default=False)

    cpf = models.CharField('CPF', max_length=14, unique=True, blank=True, null=True)
    crm = models.CharField('CRM', max_length=20, blank=True)
    uf_crm = models.CharField('UF do CRM', max_length=2, blank=True)
    especialidade = models.CharField('Especialidade', max_length=120, blank=True)
    email = models.EmailField('E-mail', blank=True)
    telefone = models.CharField('Telefone', max_length=20, blank=True)

    ativo = models.BooleanField('Ativo', default=True)
    observacoes = models.TextField('Observações', blank=True)

    criado_em = models.DateTimeField('Criado em', auto_now_add=True)
    atualizado_em = models.DateTimeField('Atualizado em', auto_now=True)

    class Meta:
        verbose_name = 'Médico'
        verbose_name_plural = 'Médicos'
        ordering = ['nome']

    def __str__(self):
        if self.crm:
            return f'{self.nome} (CRM {self.crm}/{self.uf_crm})'
        return self.nome


class CorrecaoMemorizada(models.Model):
    """Ajuste manual que o sistema memoriza para reaplicar nos próximos meses.

    É o coração da "gestão da informação": quando a diretoria corrige um valor na
    revisão e marca "memorizar", o ajuste deixa de viver na cabeça de alguém e
    passa a ser a fonte única — reaplicado automaticamente quando o mesmo
    procedimento/convênio aparecer de novo, sem ninguém precisar refazer.

    Casamento por (procedimento, convênio, médico) já normalizados. Convênio ou
    médico em branco = vale para qualquer convênio / qualquer médico; a busca
    prefere sempre a regra mais específica.
    """

    TIPO_FIXO = 'fixo'
    TIPO_PERCENTUAL = 'percentual'
    TIPO_CHOICES = [
        (TIPO_FIXO, 'Valor fixo (R$)'),
        (TIPO_PERCENTUAL, 'Percentual do valor bruto'),
    ]

    procedimento = models.CharField('Procedimento', max_length=255)
    proc_norm = models.CharField(max_length=255, editable=False, db_index=True)
    convenio = models.CharField('Convênio', max_length=120, blank=True,
                                help_text='Em branco = vale para qualquer convênio.')
    conv_norm = models.CharField(max_length=120, blank=True, editable=False, db_index=True)
    medico = models.CharField('Médico', max_length=200, blank=True,
                              help_text='Em branco = vale para qualquer médico.')
    medico_norm = models.CharField(max_length=200, blank=True, editable=False, db_index=True)

    tipo = models.CharField('Tipo', max_length=12, choices=TIPO_CHOICES, default=TIPO_FIXO)
    valor = models.DecimalField('Valor', max_digits=12, decimal_places=4,
                                help_text='Fixo: em Reais. Percentual: fração (ex.: 0,24 = 24%).')

    ativo = models.BooleanField('Ativa', default=True)
    origem = models.CharField('Origem', max_length=200, blank=True,
                              help_text='De onde veio (ex.: arquivo do lote que originou).')
    observacao = models.CharField('Observação', max_length=255, blank=True)

    criado_em = models.DateTimeField('Criada em', auto_now_add=True)
    atualizado_em = models.DateTimeField('Atualizada em', auto_now=True)

    class Meta:
        verbose_name = 'Correção memorizada'
        verbose_name_plural = 'Correções memorizadas'
        ordering = ['procedimento', 'convenio', 'medico']
        unique_together = [('proc_norm', 'conv_norm', 'medico_norm')]

    def save(self, *args, **kwargs):
        # Mantém os campos normalizados em sincronia com os textos (inclusive no admin).
        from .services.regras import normalizar
        self.proc_norm = normalizar(self.procedimento)
        self.conv_norm = normalizar(self.convenio)
        self.medico_norm = normalizar(_RE_SUFIXO.sub('', self.medico or ''))
        super().save(*args, **kwargs)

    @property
    def valor_legivel(self):
        if self.tipo == self.TIPO_PERCENTUAL:
            return f'{float(self.valor) * 100:.2f}%'.replace('.', ',') + ' do bruto'
        v = f'{float(self.valor):,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')
        return f'R$ {v}'

    def __str__(self):
        alvo = self.convenio or 'qualquer convênio'
        return f'{self.procedimento} / {alvo}'


class RepasseRascunho(models.Model):
    """Memória de curto prazo das edições de UM repasse em andamento.

    Guarda os campos da tela de revisão (honorários, classes, catarata,
    anestesista, preceptoria) ligados ao arquivo importado (token). Assim as
    alterações ficam salvas enquanto a pessoa trabalha, sem precisar refazer tudo
    de uma vez. A memória é zerada quando um novo repasse é importado."""

    token = models.CharField('Token do upload', max_length=40, unique=True, db_index=True)
    dados = models.JSONField('Edições', default=dict)
    atualizado_em = models.DateTimeField('Atualizado em', auto_now=True)

    class Meta:
        verbose_name = 'Rascunho de repasse'
        verbose_name_plural = 'Rascunhos de repasse'

    def __str__(self):
        return f'Rascunho {self.token} ({len(self.dados)} campos)'
