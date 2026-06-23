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
    classe = models.CharField('Classe', max_length=40, blank=True,
                              help_text='Se preenchida, também força a classe (ex.: laudos).')

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


class RegraRepasse(models.Model):
    """Regra de honorário por procedimento — geridas DENTRO do sistema (antes vinham
    só da planilha). Cada coluna é o valor para um pagador.

    Formato de cada valor (texto): em branco = não se aplica àquele pagador;
    "-" = não recebe (R$ 0); número = R$ fixo (ex.: "25", "1120", "549,99");
    percentual com % sobre o valor do paciente (ex.: "24%")."""

    CLASSE_CIRURGIA = 'Cirurgias e Procedimentos'
    CLASSE_EXAME = 'Exames e Consultas'
    CLASSE_PRECEPTORIA = 'Preceptoria'
    CLASSE_CHOICES = [
        (CLASSE_CIRURGIA, 'Cirurgias e Procedimentos'),
        (CLASSE_EXAME, 'Exames e Consultas'),
        (CLASSE_PRECEPTORIA, 'Preceptoria'),
    ]
    _AJUDA = 'Em branco = não se aplica. "-" = não recebe. Número = R$ (ex.: 25). Percentual com % (ex.: 24%).'

    classe = models.CharField('Classe', max_length=40, choices=CLASSE_CHOICES, default=CLASSE_EXAME)
    nome = models.CharField('Procedimento', max_length=255)
    nome_norm = models.CharField(max_length=255, editable=False, db_index=True)

    val_particular = models.CharField('Particular', max_length=20, blank=True, help_text=_AJUDA)
    val_convenio = models.CharField('Convênio', max_length=20, blank=True, help_text=_AJUDA)
    val_sus = models.CharField('SUS', max_length=20, blank=True, help_text=_AJUDA)
    val_oci = models.CharField('OCI', max_length=20, blank=True, help_text=_AJUDA)
    val_cisa = models.CharField('CISA', max_length=20, blank=True, help_text=_AJUDA)

    ativo = models.BooleanField('Ativa', default=True)
    observacao = models.CharField('Observação', max_length=255, blank=True)
    criado_em = models.DateTimeField('Criada em', auto_now_add=True)
    atualizado_em = models.DateTimeField('Atualizada em', auto_now=True)

    class Meta:
        verbose_name = 'Regra de repasse'
        verbose_name_plural = 'Regras de repasse'
        ordering = ['classe', 'nome']
        unique_together = [('nome_norm', 'classe')]

    def save(self, *args, **kwargs):
        from .services.regras import normalizar
        self.nome_norm = normalizar(self.nome)
        super().save(*args, **kwargs)

    def __str__(self):
        return f'{self.nome} [{self.classe}]'


class RepasseRascunho(models.Model):
    """Memória de curto prazo das edições de UM repasse em andamento.

    Guarda os campos da tela de revisão (honorários, classes, catarata,
    anestesista, preceptoria) ligados ao arquivo importado (token). Assim as
    alterações ficam salvas enquanto a pessoa trabalha, sem precisar refazer tudo
    de uma vez. A memória é zerada quando um novo repasse é importado."""

    token = models.CharField('Token do upload', max_length=40, unique=True, db_index=True)
    arquivo_nome = models.CharField('Arquivo importado', max_length=255, blank=True)
    dados = models.JSONField('Edições', default=dict)
    atualizado_em = models.DateTimeField('Atualizado em', auto_now=True)

    class Meta:
        verbose_name = 'Rascunho de repasse'
        verbose_name_plural = 'Rascunhos de repasse'

    def __str__(self):
        return f'Rascunho {self.token} ({len(self.dados)} campos)'


class Lote(models.Model):
    """Registro de um processamento exportado — o histórico/auditoria do sistema.

    Cada exportação vira (ou atualiza) um lote ligado ao arquivo importado: guarda
    quando, quem, o período, os totais, os arquivos gerados (para re-baixar sem
    subir de novo), o que foi ajustado manualmente (auditoria) e as "impressões
    digitais" dos atendimentos (para avisar se algo já saiu num lote anterior)."""

    token = models.CharField('Token do upload', max_length=40, unique=True, db_index=True)
    criado_em = models.DateTimeField('Criado em', auto_now_add=True)
    atualizado_em = models.DateTimeField('Atualizado em', auto_now=True)
    criado_por = models.CharField('Por', max_length=150, blank=True)

    arquivo_nome = models.CharField('Arquivo importado', max_length=255, blank=True)
    unidade = models.CharField('Unidade', max_length=200, blank=True)
    periodo_inicio = models.DateField('Início', null=True, blank=True)
    periodo_fim = models.DateField('Fim', null=True, blank=True)

    n_medicos = models.PositiveIntegerField('Médicos', default=0)
    n_repasses = models.PositiveIntegerField('Repasses', default=0)
    total_pagar = models.DecimalField('Total a pagar', max_digits=12, decimal_places=2, default=0)
    total_receber = models.DecimalField('Total a receber', max_digits=12, decimal_places=2, default=0)

    pasta_saida = models.CharField('Pasta de saída', max_length=120, blank=True)
    downloads = models.JSONField('Arquivos gerados', default=list)   # [{grupo, arquivo}]
    auditoria = models.JSONField('Ajustes manuais', default=list)    # lista de textos
    fingerprints = models.JSONField('Atendimentos', default=list)    # p/ duplicidade
    edicoes = models.JSONField('Edições da revisão', default=dict)   # rascunho p/ reabrir e editar
    upload_conteudo = models.BinaryField('Relatório importado', null=True, blank=True)  # o .xls de origem

    class Meta:
        verbose_name = 'Lote (histórico)'
        verbose_name_plural = 'Lotes (histórico)'
        ordering = ['-criado_em']

    def __str__(self):
        return f'Lote {self.id} — {self.arquivo_nome or self.token} ({self.criado_em:%d/%m/%Y})'

    @property
    def fixado(self):
        """Lote 'fixado' = tem algum repasse já PAGO. Fixado não pode ser editado/excluído
        (a diretoria reverte o status do pagamento para liberar)."""
        return self.repasses.filter(status=Repasse.STATUS_PAGO).exists()


class ArquivoSaida(models.Model):
    """Bytes de um arquivo gerado, guardados NO BANCO para re-download mesmo que a
    pasta de saídas seja limpa — recuperabilidade do histórico (auditoria)."""

    lote = models.ForeignKey(Lote, on_delete=models.CASCADE, related_name='arquivos')
    grupo = models.CharField('Grupo', max_length=120, blank=True)
    nome = models.CharField('Arquivo', max_length=200)
    conteudo = models.BinaryField('Conteúdo')

    class Meta:
        verbose_name = 'Arquivo gerado'
        verbose_name_plural = 'Arquivos gerados'
        ordering = ['id']

    def __str__(self):
        return self.nome


class Repasse(models.Model):
    """Um repasse individual (médico/dia/clínica) com seu andamento de pagamento.

    Permite acompanhar cada repasse de 'gerado' até 'pago' — o painel do que
    ainda falta enviar/pagar. Mantém o status mesmo se o lote for re-exportado."""

    STATUS_GERADO = 'gerado'
    STATUS_REVISADO = 'revisado'
    STATUS_ENVIADO = 'enviado'
    STATUS_PAGO = 'pago'
    STATUS_CHOICES = [
        (STATUS_GERADO, 'Gerado'),
        (STATUS_REVISADO, 'Revisado'),
        (STATUS_ENVIADO, 'Enviado ao médico'),
        (STATUS_PAGO, 'Pago'),
    ]

    lote = models.ForeignKey(Lote, on_delete=models.CASCADE, related_name='repasses')
    tipo = models.CharField(max_length=12, default='medico')  # medico | anestesista
    medico = models.CharField('Profissional', max_length=200)
    razao_social = models.CharField('Razão Social', max_length=200, blank=True)
    data = models.DateField('Data', null=True, blank=True)
    clinica = models.CharField('Clínica', max_length=200, blank=True)
    valor = models.DecimalField('Valor', max_digits=12, decimal_places=2, default=0)
    status = models.CharField('Status', max_length=12, choices=STATUS_CHOICES, default=STATUS_GERADO)
    atualizado_em = models.DateTimeField('Atualizado em', auto_now=True)

    class Meta:
        verbose_name = 'Repasse (acompanhamento)'
        verbose_name_plural = 'Repasses (acompanhamento)'
        ordering = ['data', 'clinica', 'medico']
        unique_together = [('lote', 'tipo', 'medico', 'data', 'clinica')]

    @property
    def pago(self):
        return self.status == self.STATUS_PAGO

    def __str__(self):
        dia = self.data.strftime('%d/%m') if self.data else 's/ data'
        return f'{self.medico} {dia} {self.clinica} — {self.get_status_display()}'
