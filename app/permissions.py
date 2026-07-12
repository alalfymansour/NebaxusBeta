from functools import wraps
from flask import session, redirect, url_for, flash, request, jsonify, g

CUSTOMER_PERMISSIONS = [
    'can_view_customers', 'can_add_customers', 'can_edit_customers',
    'can_delete_customers', 'can_view_customer_orders', 'can_view_customer_logs',
    'can_view_customer_profile', 'can_add_customer_logs', 'can_delete_customer_logs'
]

PRODUCT_PERMISSIONS = [
    'can_view_products', 'can_add_products', 'can_edit_products',
    'can_delete_products', 'can_restore_products', 'can_delete_all_products',
    'can_view_purchase_price'
]

ORDER_PERMISSIONS = [
    'can_view_orders', 'can_add_orders', 'can_edit_orders', 'can_delete_orders',
    'can_update_order_status', 'can_update_tracking', 'can_view_orders_by_status'
]

REPLACEMENT_PERMISSIONS = [
    'can_view_replacements', 'can_add_replacements', 'can_edit_replacements',
    'can_view_replacements_by_state', 'can_delete_replacements'
]

RETURN_PERMISSIONS = [
    'can_view_returns',
    'can_add_returns',
    'can_view_returns_by_state',
    'can_delete_returns'
]

FEES_PERMISSIONS = ['can_view_fees', 'can_edit_fees', 'can_delete_fees']

EMPLOYEE_PERMISSIONS = [
    'can_view_employees', 'can_add_employees', 'can_edit_employees', 'can_delete_employees',
    'can_view_employee_logs', 'can_manage_employee_salary', 'can_delete_salary_transactions',
    'can_view_employee_activity'
]

REPORTS_PERMISSIONS = [
    'can_view_statistics',
    'can_view_stats_stock', 'can_view_stats_fixed_assets', 'can_view_stats_total_debt',
    'can_view_stats_capital_growth', 'can_view_stats_daily', 'can_view_stats_pending_orders',
    'can_view_stats_net_profit', 'can_view_stats_losses', 'can_view_stats_delivery_rate',
    'can_view_stats_sales', 'can_view_stats_amount_paid', 'can_view_stats_monthly_delivered',
    'can_view_stats_monthly_invoices', 'can_view_stats_fixed_assets_expenses',
    'can_view_stats_employee_salaries', 'can_view_stats_operational_expenses',
    'can_view_sold_products_by_quantity'
]

SUPPLIER_PERMISSIONS = [
    'can_view_suppliers', 'can_add_suppliers', 'can_edit_suppliers',
    'can_delete_suppliers', 'can_view_supplier_invoices', 'can_view_supplier_history',
    'can_pay_supplier_debt', 'can_add_supplier_debt'
]

INVOICE_PERMISSIONS = [
    'can_view_invoices', 'can_add_invoices', 'can_edit_invoices', 'can_delete_invoices',
    'can_view_invoice_details'
]

EXPENSE_PERMISSIONS = [
    'can_view_expenses', 'can_add_expenses', 'can_edit_expenses', 'can_delete_expenses',
    'can_view_operational_expenses', 'can_view_fixed_assets'
]

SYSTEM_PERMISSIONS = [
    'can_view_activity_log', 'can_manage_reminders'
]

FOLLOWUP_PERMISSIONS = [
    'can_view_followups', 'can_add_followups', 'can_edit_followups', 'can_delete_followups'
]

DAMAGED_PRODUCTS_PERMISSIONS = [
    'can_view_damaged_products', 'can_add_damaged_products', 'can_delete_damaged_products'
]

CARD_TRANSACTIONS_PERMISSIONS = [
    'can_view_transactions',
    'can_add_transactions',
    'can_edit_transactions',
    'can_delete_transactions'
]

ALL_PERMISSIONS = (
    CUSTOMER_PERMISSIONS + PRODUCT_PERMISSIONS + ORDER_PERMISSIONS +     REPLACEMENT_PERMISSIONS + RETURN_PERMISSIONS +
    FEES_PERMISSIONS + EMPLOYEE_PERMISSIONS + REPORTS_PERMISSIONS + SUPPLIER_PERMISSIONS +
    INVOICE_PERMISSIONS + EXPENSE_PERMISSIONS + SYSTEM_PERMISSIONS + FOLLOWUP_PERMISSIONS +
    DAMAGED_PRODUCTS_PERMISSIONS + CARD_TRANSACTIONS_PERMISSIONS
)

def build_permission_dict(employee) -> dict:
    data = {}
    for perm in ALL_PERMISSIONS:
        data[perm] = getattr(employee, perm, False)
    data['is_admin'] = getattr(employee, 'is_admin', False)
    data['is_active'] = getattr(employee, 'is_active', True)
    return data

def _get_current_employee_permissions() -> dict:
    if getattr(g, '_employee_perms_cache', None) is not None:
        return g._employee_perms_cache

    perms = {}
    try:
        emp_id = session.get('employee_id')
        if not emp_id:
            g._employee_perms_cache = perms
            return perms
        from .models import Employee
        from . import db
        employee = db.session.get(Employee, emp_id)
        if not employee:
            g._employee_perms_cache = perms
            return perms
        for p in ALL_PERMISSIONS:
            perms[p] = getattr(employee, p, False)
        perms['is_admin'] = getattr(employee, 'is_admin', False)
        perms['is_active'] = getattr(employee, 'is_active', True)
    except Exception:
        perms = {}
    g._employee_perms_cache = perms
    return perms

def require_permissions(*permissions, any_: bool = False, admin_override: bool = True):
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if 'employee_id' not in session:
                flash('يجب تسجيل الدخول أولاً')
                return redirect(url_for('main.login'))

            # Always read is_admin fresh from DB (never trust stale session value)
            perms = _get_current_employee_permissions()

            if admin_override and perms.get('is_admin'):
                return f(*args, **kwargs)

            if not permissions:
                return f(*args, **kwargs)

            if any_:
                ok = any(perms.get(p) for p in permissions)
            else:
                ok = all(perms.get(p) for p in permissions)
            if not ok:
                message = 'ليس لديك هذه الصلاحية'
                want_json = request.is_json or request.accept_mimetypes.best == 'application/json' or request.path.startswith('/api/')
                if want_json:
                    return jsonify({'error': message, 'status': 403}), 403
                flash(message)
                return redirect(url_for('main.dashboard'))
            return f(*args, **kwargs)
        return wrapper
    return decorator

def has_permission(perm: str) -> bool:
    # Always read is_admin fresh from DB (never trust stale session value)
    perms = _get_current_employee_permissions()
    if perms.get('is_admin'):
        return True
    return perms.get(perm, False)

def has_any(*perms) -> bool:
    return any(has_permission(p) for p in perms)

def has_all(*perms) -> bool:
    return all(has_permission(p) for p in perms)

__all__ = [
    'ALL_PERMISSIONS', 'build_permission_dict', 'require_permissions',
    'has_permission', 'has_any', 'has_all'
]
