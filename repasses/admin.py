from django.contrib import admin

from .models import CorrecaoMemorizada, Lote, Medico, Repasse


@admin.register(Medico)
class MedicoAdmin(admin.ModelAdmin):
    list_display = ('nome', 'categoria', 'eh_fellow', 'eh_anestesista', 'razao_social', 'ativo')
    list_filter = ('categoria', 'eh_fellow', 'eh_preceptor', 'eh_anestesista', 'ativo')
    search_fields = ('nome', 'cpf', 'crm', 'email', 'razao_social')
    readonly_fields = ('criado_em', 'atualizado_em')
    fieldsets = (
        ('Identificação', {
            'fields': ('nome', 'categoria', ('eh_fellow', 'eh_preceptor', 'eh_anestesista'),
                       'razao_social', 'regra_obs'),
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


@admin.register(Repasse)
class RepasseAdmin(admin.ModelAdmin):
    list_display = ('medico', 'data', 'clinica', 'tipo', 'valor', 'status', 'lote')
    list_filter = ('status', 'tipo', 'clinica')
    search_fields = ('medico', 'razao_social', 'clinica')
    list_editable = ('status',)
