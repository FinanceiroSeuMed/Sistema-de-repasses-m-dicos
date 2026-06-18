from django.contrib import admin

from .models import Medico


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
