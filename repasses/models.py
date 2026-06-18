"""
Modelos de dados do sistema de repasses médicos.

Por enquanto temos o cadastro de Médicos, que é a base de qualquer repasse.
Os demais modelos (procedimentos, regras de honorário, repasses gerados,
integrações com MedPlus e OMIE) serão adicionados conforme recebermos os
modelos de relatório e de importação.
"""

from django.db import models


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
