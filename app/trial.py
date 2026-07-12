import json
import zipfile
from datetime import datetime
from io import BytesIO

from flask import jsonify, request, redirect, url_for, flash

from app import db
from app.models import Order, Customer, Product, Employee

TRIAL_LIMIT = 500


def orders_this_month_count():
    now = datetime.utcnow()
    first = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    return Order.query.filter(Order.date >= first).count()


def remaining():
    return max(0, TRIAL_LIMIT - orders_this_month_count())


def is_read_only():
    return remaining() <= 0


def inject_trial_context():
    used = orders_this_month_count()
    return {
        'trial_limit': TRIAL_LIMIT,
        'trial_used': used,
        'trial_remaining': max(0, TRIAL_LIMIT - used),
        'trial_read_only': used >= TRIAL_LIMIT,
        'trial_pct': min(100, int(used / TRIAL_LIMIT * 100)),
    }


CHECK_ENDPOINTS = {
    'main.add_order',
    'main.add_return_order',
    'main.add_replacement_order',
}


def before_request_check():
    if request.method != 'POST':
        return None
    if request.endpoint not in CHECK_ENDPOINTS:
        return None
    if not is_read_only():
        return None
    if request.is_json or request.accept_mimetypes.accept_json:
        return jsonify({
            'success': False,
            'error': 'تم الوصول للحد الأقصى للطلبات التجريبية (500 طلب/شهر). قم بالترقية على nebaxus.com'
        }), 403
    flash(
        'تم الوصول للحد الأقصى للطلبات التجريبية (500 طلب/شهر). قم بالترقية للإصدار الكامل على nebaxus.com'
    )
    return redirect(url_for('main.dashboard'))


def generate_nbx():
    now = datetime.utcnow()
    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        manifest = {
            'format': 'nbx',
            'version': 1,
            'created_at': now.isoformat(),
            'application': 'NebaxusBeta',
        }
        zf.writestr('manifest.json', json.dumps(manifest, ensure_ascii=False, indent=2))

        tables = {
            'customers': [{
                'id': c.id, 'name': c.name, 'phone': c.phone,
                'address': c.address, 'governorate': c.governorate,
                'created_at': c.created_at.isoformat() if c.created_at else None,
            } for c in Customer.query.all()],
            'orders': [{
                'id': o.id, 'customer_id': o.customer_id,
                'total': float(o.total or 0), 'status': o.status,
                'created_at': o.created_at.isoformat() if o.created_at else None,
            } for o in Order.query.all()],
            'products': [{
                'id': p.id, 'name': p.name,
                'price': float(p.price or 0), 'stock': p.stock,
            } for p in Product.query.all()],
            'employees': [{
                'id': e.id, 'name': e.name, 'username': e.username,
                'phone': e.phone, 'is_admin': e.is_admin,
            } for e in Employee.query.all()],
        }
        for name, rows in tables.items():
            data = json.dumps(rows, ensure_ascii=False, indent=2, default=str)
            zf.writestr(f'{name}.json', data)

    buf.seek(0)
    return buf
