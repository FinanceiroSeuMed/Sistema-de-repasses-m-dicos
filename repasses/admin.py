from django.contrib import admin

from .models import Medico


@admin.register(Medico)
class MedicoAdmin(admin.ModelAdmin):
    list_display = ('nome', 'crm', 'uf_crm', 'especialidade', 'ativo', 'atualizado_em')
    list_filter = ('ativo', 'uf_crm', 'especialidade')
    search_fields = ('nome', 'cpf', 'crm', 'email')
    readonly_fields = ('criado_em', 'atualizado_em')
    fieldsets = (
        ('Identificação', {
            'fields': ('nome', 'cpf', ('crm', 'uf_crm'), 'especialidade'),
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
