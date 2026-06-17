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

    nome = models.CharField('Nome completo', max_length=200)
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
