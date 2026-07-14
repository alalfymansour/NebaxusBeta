_MODELS: dict[str, type] = {}

def _model(name: str) -> type:
    if not _MODELS:
        from .models import Customer, Order, Product, Employee, Supplier, Invoice, Expense, ReplacementOrder, FollowUp, Party  # noqa: F401
        _MODELS.update({k: v for k, v in locals().items() if k[0].isupper()})
    return _MODELS[name]


class _Dyn:
    def __init__(self, model_name, id_param, label_prefix='', icon=None):
        self.model_name = model_name
        self.id_param = id_param
        self.label_prefix = label_prefix
        self.icon = icon

    def resolve(self, **kwargs):
        entity_id = kwargs.get(self.id_param)
        entity = _model(self.model_name).query.get(entity_id) if entity_id else None
        if entity:
            for attr in ('name', 'title', 'invoice_number', 'description'):
                val = getattr(entity, attr, None)
                if val:
                    name = val
                    break
            else:
                name = f'#{entity.id}'
            label = f'{self.label_prefix}: {name}' if self.label_prefix else name
        else:
            label = self.label_prefix or '—'
        return (label, self.icon)


class _DynStr:
    """Resolves label from URL param (e.g. orders_by_status(status='جديد'))."""
    def __init__(self, prefix, param, icon=None):
        self.prefix = prefix
        self.param = param
        self.icon = icon

    def resolve(self, **kwargs):
        val = kwargs.get(self.param, '')
        label = f'{self.prefix} - {val}' if val else self.prefix
        return (label, self.icon)


def page_label(endpoint, **kwargs):
    func_name = endpoint.rsplit('.', 1)[-1]
    spec = BREADCRUMBS.get(func_name)
    if spec is None:
        return (None, None)
    if isinstance(spec, (_Dyn, _DynStr)):
        return spec.resolve(**kwargs)
    return spec


BREADCRUMBS = {
    # Dashboard
    'dashboard': ('الرئيسية', 'house'),

    # Orders
    'orders': ('الطلبات', 'receipt'),
    'orders_by_status': _DynStr('الطلبات', 'status', icon='receipt'),
    'add_order': ('إضافة طلب', 'plus-circle'),
    'edit_order': _Dyn('Order', 'order_id', 'تعديل', icon='pencil'),
    'order_status_history': _Dyn('Order', 'order_id', 'سجل الحالة', icon='clock'),
    'new_products_summary': ('ملخص المنتجات الجديدة', 'stack'),
    'new_products_summary_page': ('ملخص المنتجات الجديدة', 'stack'),
    'orders_with_missing_products': ('طلبات بمنتجات مفقودة', 'warning-circle'),

    # Returns
    'returns': ('المرتجعات', 'arrow-u-up-left'),

    'add_return_order': ('إضافة مرتجع', 'plus-circle'),

    # Replacements
    'replacement_orders': ('طلبات الاستبدال', 'arrows-left-right'),
    'replacement_orders_by_status': _DynStr('طلبات الاستبدال', 'status', icon='arrows-left-right'),
    'add_replacement_order': ('إضافة طلب استبدال', 'plus-circle'),
    'edit_replacement_order': _Dyn('ReplacementOrder', 'order_id', 'تعديل', icon='pencil'),
    'replacement_drafts': ('مسودات الاستبدال', 'file-dashed'),
    'edit_replacement_draft': _Dyn('ReplacementOrder', 'draft_id', 'تعديل مسودة', icon='pencil'),
    'replacement_order_status_history': _Dyn('ReplacementOrder', 'order_id', 'سجل الحالة', icon='clock'),
    'replacement_order_losses': ('خسائر الاستبدال', 'trend-down'),
    'replacements_shipped': ('طلبات تم شحنها', 'truck'),

    # Customers
    'customers': ('العملاء', 'users'),
    'customers_more': ('العملاء', 'users'),
    'add_customer': ('إضافة عميل', 'plus-circle'),
    'edit_customer': _Dyn('Customer', 'customer_id', 'تعديل', icon='pencil'),
    'customer_profile': _Dyn('Customer', 'customer_id', 'ملف', icon='user'),
    'view_customer_orders': _Dyn('Customer', 'customer_id', 'الطلبات', icon='receipt'),
    'customer_history': _Dyn('Customer', 'customer_id', 'سجل', icon='clock'),
    'add_customer_log': _Dyn('Customer', 'customer_id', 'إضافة ملاحظة', icon='note-pencil'),

    # Products
    'products': ('المنتجات', 'stack'),
    'add_product': ('إضافة منتج', 'plus-circle'),
    'edit_product': _Dyn('Product', 'product_id', 'تعديل', icon='pencil'),
    'damaged_products': ('المنتجات التالفة', 'warning-circle'),
    'add_damaged_product': ('إضافة منتج تالف', 'plus-circle'),

    # Stocktake
    'stocktake_index': ('المخازن', 'warehouse'),
    'stocktake_list': ('سجل الجرد', 'clipboard-text'),
    'stocktake_new': ('جرد جديد', 'plus-circle'),
    'stocktake_summary': ('ملخص الجرد', 'clipboard-text'),

    # Suppliers
    'suppliers': ('الموردين', 'handshake'),
    'add_supplier': ('إضافة مورد', 'plus-circle'),
    'edit_supplier': _Dyn('Supplier', 'supplier_id', 'تعديل', icon='pencil'),
    'supplier_profile': _Dyn('Supplier', 'supplier_id', 'ملف', icon='user'),
    'supplier_invoices': _Dyn('Supplier', 'supplier_id', 'الفواتير', icon='file-text'),
    'supplier_account_history': _Dyn('Supplier', 'supplier_id', 'سجل الحساب', icon='book'),

    # Employees
    'employees': ('الموظفين', 'user-gear'),
    'add_employee': ('إضافة موظف', 'plus-circle'),
    'add_employee_basic': ('إضافة موظف - بيانات', 'plus-circle'),
    'add_employee_permissions': ('إضافة موظف - الصلاحيات', 'gear'),
    'edit_employee': _Dyn('Employee', 'employee_id', 'تعديل', icon='pencil'),
    'employee_profile': _Dyn('Employee', 'employee_id', 'ملف', icon='user'),
    'employee_salary': _Dyn('Employee', 'employee_id', 'الراتب', icon='currency-dollar'),
    'employee_permissions': _Dyn('Employee', 'employee_id', 'الصلاحيات', icon='gear'),
    'employee_activity_log': _Dyn('Employee', 'employee_id', 'النشاط', icon='clock'),

    # Invoices
    'invoices': ('الفواتير', 'file-text'),
    'add_invoice': ('إضافة فاتورة', 'plus-circle'),
    'edit_invoice': _Dyn('Invoice', 'invoice_id', 'تعديل', icon='pencil'),
    'view_invoice': _Dyn('Invoice', 'invoice_id', 'عرض', icon='file-text'),

    # Expenses
    'expenses': ('المصروفات', 'calculator'),
    'add_expense': ('إضافة مصروف', 'plus-circle'),
    'view_expense': _Dyn('Expense', 'expense_id', 'عرض', icon='calculator'),
    'operational_expenses': ('مصروفات تشغيلية', 'calculator'),
    'fixed_assets_expenses': ('أصول ثابتة', 'buildings'),

    # Transactions
    'transactions': ('المعاملات', 'credit-card'),
    'transactions_more': ('المعاملات', 'credit-card'),
    'party_profile': _Dyn('Party', 'party_id', 'ملف', icon='user'),
    'add_party': ('إضافة متعامل', 'plus-circle'),
    'edit_party': _Dyn('Party', 'party_id', 'تعديل', icon='pencil'),
    'add_transaction_page': ('إضافة معاملة', 'plus-circle'),
    'add_transaction': ('إضافة معاملة', 'plus-circle'),

    # Attendance
    'attendance': _Dyn('Employee', 'employee_id', 'الحضور', icon='calendar-check'),
    'attendance_today': ('حضور اليوم', 'calendar-check'),

    # Followups
    'followups': ('قائمة المتابعة', 'clipboard-text'),
    'add_followup': ('إضافة متابعة', 'plus-circle'),

    # Statistics
    'statistics': ('الإحصائيات', 'chart-bar'),

    # Misc
    'activity_log': ('سجل النشاطات', 'clock'),
    'my_account': ('حسابي', 'user'),
    'governorate_fees': ('رسوم المحافظات', 'currency-dollar'),
}
