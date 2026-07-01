from django.contrib import admin

from .models import (AjusteMensal, ClasseMemorizada, CorrecaoMemorizada, Lote, Medico,
                     RegraRepasse, Repasse)


@admin.register(AjusteMensal)
class AjusteMensalAdmin(admin.ModelAdmin):
    list_display = ('ano_mes', 'medico', 'valor', 'motivo', 'atualizado_em')
    list_filter = ('ano_mes',)
    search_fields = ('medico__nome', 'motivo')
    readonly_fields = ('criado_em', 'atualizado_em')


@admin.register(Medico)
class MedicoAdmin(admin.ModelAdmin):
    list_display = ('nome', 'categoria', 'eh_fellow', 'eh_anestesista', 'razao_social',
                    'cnpj', 'chave_pix', 'ativo')
    list_filter = ('categoria', 'eh_fellow', 'eh_preceptor', 'eh_anestesista', 'ativo')
    search_fields = ('nome', 'cpf', 'crm', 'email', 'razao_social', 'cnpj', 'chave_pix')
    readonly_fields = ('criado_em', 'atualizado_em')
    fieldsets = (
        ('Identificação', {
            'fields': ('nome', 'categoria', ('eh_fellow', 'eh_preceptor', 'eh_anestesista'),
                       'razao_social', ('cnpj', 'chave_pix'), 'regra_obs'),
        }),
        ('Documentos', {
            'fields': ('cpf', ('crm', 'uf_crm'), 'especialidade'),
        }),
        ('Contato', {
            'fields': ('email', 'telefone'),
        }),
        ('Situação', {
            'fields': ('ativo', 'observacoes'),
        }),
        ('Auditoria', {
            'fields': ('criado_em', 'atualizado_em'),
            'classes': ('collapse',),
        }),
    )


@admin.register(CorrecaoMemorizada)
class CorrecaoMemorizadaAdmin(admin.ModelAdmin):
    list_display = ('procedimento', 'convenio', 'medico', 'tipo', 'valor', 'ativo', 'atualizado_em')
    list_filter = ('ativo', 'tipo')
    search_fields = ('procedimento', 'convenio', 'medico', 'origem', 'observacao')
    readonly_fields = ('proc_norm', 'conv_norm', 'medico_norm', 'criado_em', 'atualizado_em')
    fields = ('procedimento', 'convenio', 'medico', ('tipo', 'valor'), 'ativo',
              'origem', 'observacao', ('proc_norm', 'conv_norm', 'medico_norm'),
              ('criado_em', 'atualizado_em'))


@admin.register(ClasseMemorizada)
class ClasseMemorizadaAdmin(admin.ModelAdmin):
    list_display = ('procedimento', 'classe', 'ativo', 'atualizado_em')
    list_filter = ('ativo', 'classe')
    search_fields = ('procedimento', 'origem')
    readonly_fields = ('proc_norm', 'criado_em', 'atualizado_em')
    fields = ('procedimento', 'classe', 'ativo', 'origem', 'proc_norm',
              ('criado_em', 'atualizado_em'))


@admin.register(Lote)
class LoteAdmin(admin.ModelAdmin):
    list_display = ('id', 'criado_em', 'arquivo_nome', 'unidade', 'n_medicos',
                    'total_pagar', 'total_receber', 'criado_por')
    list_filter = ('unidade', 'criado_por')
    search_fields = ('arquivo_nome', 'unidade', 'token')
    readonly_fields = ('token', 'criado_em', 'atualizado_em', 'criado_por', 'arquivo_nome',
                       'unidade', 'periodo_inicio', 'periodo_fim', 'n_medicos', 'n_repasses',
                       'total_pagar', 'total_receber', 'pasta_saida', 'downloads',
                       'auditoria', 'fingerprints')


@admin.register(RegraRepasse)
class RegraRepasseAdmin(admin.ModelAdmin):
    list_display = ('nome', 'classe', 'val_particular', 'val_convenio', 'val_sus',
                    'val_oci', 'val_cisa', 'ativo')
    list_editable = ('val_particular', 'val_convenio', 'val_sus', 'val_oci', 'val_cisa', 'ativo')
    list_filter = ('classe', 'ativo')
    search_fields = ('nome', 'observacao')
    readonly_fields = ('nome_norm', 'criado_em', 'atualizado_em')
    list_per_page = 200


@admin.register(Repasse)
class RepasseAdmin(admin.ModelAdmin):
    list_display = ('medico', 'data', 'clinica', 'tipo', 'valor', 'status', 'lote')
    list_filter = ('status', 'tipo', 'clinica')
    search_fields = ('medico', 'razao_social', 'clinica')
    list_editable = ('status',)
