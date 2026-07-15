import json
from . import db  
from flask_login import UserMixin
from datetime import datetime, date
from calendar import monthrange
from sqlalchemy import event
from sqlalchemy import inspect as sa_inspect
from sqlalchemy.orm import backref

"""
Dukana ORM models.

All SQLAlchemy model classes for the application are defined here.
Models are grouped in this order:
  1. Customer & product domain  (Customer, Product, ProductVariant, BundleItem)
  2. Order domain               (Order, OrderItem, OrderStatusHistory)
  3. Replacement order domain   (ReplacementOrder, ReplacementOrderItem, ...)
  4. Employee & HR domain       (Employee, SalaryTransaction, EmployeeLoginLog, ...)
  5. Supplier & finance domain  (Supplier, Invoice, Expense, ...)
  6. Reporting & settings       (CapitalSnapshot, CapitalGrowthHistory, AppSettings, ...)

NOTE: Business logic in this file is intentionally limited to model-level
computations (properties, aggregations). Route-level orchestration belongs
in routes.py or a dedicated service layer.
"""


def _parse_bundle_variants_map(bundle_variants_json: str | None) -> dict:
    """
    Parse a bundle variants JSON string into a structured dict of variant objects.

    Args:
        bundle_variants_json: Raw JSON string stored on the order item, or None.

    Returns:
        Dict keyed by item index (int) with structure:
            { item_idx: { 'product_id': int, 'variants': { 'مقاس': ProductVariant, ... } } }
        Returns empty dict if input is None, empty, or unparseable.
    """
    if not bundle_variants_json:
        return {}

    try:
        bundle_data = json.loads(bundle_variants_json)
        if not isinstance(bundle_data, dict):
            return {}

        variant_ids = set()
        for item_data in bundle_data.values():
            if not isinstance(item_data, dict):
                continue
            for key, val in item_data.items():
                if key.endswith('_variant_id') and isinstance(val, int):
                    variant_ids.add(val)

        variants_dict = {}
        if variant_ids:
            variants = db.session.query(ProductVariant).filter(ProductVariant.id.in_(variant_ids)).all()
            variants_dict = {variant.id: variant for variant in variants}

        result = {}
        for item_idx, item_data in bundle_data.items():
            if not isinstance(item_data, dict):
                continue

            product_id = item_data.get('product_id')
            if not product_id:
                continue

            try:
                parsed_item_idx = int(item_idx)
            except (TypeError, ValueError):
                continue

            result[parsed_item_idx] = {
                'product_id': product_id,
                'variants': {}
            }

            if 'size_variant_id' in item_data and item_data['size_variant_id'] in variants_dict:
                result[parsed_item_idx]['variants']['مقاس'] = variants_dict[item_data['size_variant_id']]

            if 'color_variant_id' in item_data and item_data['color_variant_id'] in variants_dict:
                result[parsed_item_idx]['variants']['لون'] = variants_dict[item_data['color_variant_id']]

            if 'style_variant_id' in item_data and item_data['style_variant_id'] in variants_dict:
                result[parsed_item_idx]['variants']['شكل'] = variants_dict[item_data['style_variant_id']]

        return result
    except (json.JSONDecodeError, KeyError, TypeError) as e:
        # Corrupted or unexpected bundle JSON format — return empty rather than crash.
        import logging
        logging.getLogger(__name__).warning(
            "Failed to parse bundle_variants_json: %s", e
        )
        return {}

class Customer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    phone = db.Column(db.String(20), nullable=False, index=True)
    governorate = db.Column(db.String(100), nullable=True)
    address_details = db.Column(db.String(200), nullable=True)

    orders = db.relationship('Order', backref='customer', lazy=True, cascade="all, delete-orphan")
    logs = db.relationship('CustomerLog', backref='customer', lazy=True, cascade="all, delete-orphan")
    followups = db.relationship('FollowUp', backref='customer', lazy=True, cascade="all, delete-orphan")


class Product(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, index=True)
    price = db.Column(db.Float, nullable=False)
    purchase_price = db.Column(db.Float, nullable=False, default=0)
    wholesale_price = db.Column(db.Float, nullable=False, default=0)
    stock = db.Column(db.Integer, nullable=False, default=0)
    is_deleted = db.Column(db.Boolean, default=False, index=True)
    
    is_bundle = db.Column(db.Boolean, default=False, nullable=False)
    
    has_size = db.Column(db.Boolean, default=False)
    has_color = db.Column(db.Boolean, default=False)
    has_style = db.Column(db.Boolean, default=False)

    variants = db.relationship('ProductVariant', backref='product', lazy=True, cascade="all, delete-orphan")
    invoice_items = db.relationship('InvoiceItem', backref='product', lazy=True, cascade="all, delete-orphan")
    issue_items = db.relationship('IssueItem', backref='product', lazy=True, cascade="all, delete-orphan")
    damaged_logs = db.relationship('DamagedProductLog', backref='product', lazy=True, cascade="all, delete-orphan")
    bundle_items = db.relationship('BundleItem', backref='bundle_product', lazy=True, cascade="all, delete-orphan", foreign_keys='BundleItem.bundle_id')
    
    @property
    def variant_types(self):
        types = []
        if self.has_size:
            types.append('مقاس')
        if self.has_color:
            types.append('لون')
        if self.has_style:
            types.append('شكل')
        return types
    
    @property
    def has_variants(self):
        return self.has_size or self.has_color or self.has_style
    
    def update_variant_types(self):
        variants = ProductVariant.query.filter_by(product_id=self.id).all()
        
        self.has_size = any(v.group_name == 'مقاس' for v in variants)
        self.has_color = any(v.group_name == 'لون' for v in variants)
        self.has_style = any(v.group_name == 'شكل' for v in variants)
        
        return self
    
    @property
    def profit(self):
        if self.is_bundle:
            return sum(item.profit for item in self.bundle_items)
        return self.price - self.purchase_price
    
    @property
    def profit_percentage(self):
        if self.is_bundle:
            total_purchase = sum(item.product.purchase_price for item in self.bundle_items if item.product)
            if total_purchase > 0:
                return (self.profit / total_purchase) * 100
            return 0
        if self.purchase_price > 0:
            return ((self.price - self.purchase_price) / self.purchase_price) * 100
        return 0
    
    @property
    def bundle_total_price(self):
        if self.is_bundle:
            return sum(item.sale_price_in_bundle for item in self.bundle_items)
        return self.price

class ProductVariant(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    group_name = db.Column(db.String(50), nullable=False)
    variant_name = db.Column(db.String(100), nullable=False)
    price = db.Column(db.Float, nullable=False, default=0)
    stock = db.Column(db.Integer, default=0)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id', ondelete='CASCADE', name='fk_productvariant_product_id'), nullable=False)


class BundleItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    bundle_id = db.Column(db.Integer, db.ForeignKey('product.id', ondelete='CASCADE'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id', ondelete='CASCADE'), nullable=False)
    sale_price_in_bundle = db.Column(db.Float, nullable=False, default=0)
    
    product = db.relationship('Product', foreign_keys=[product_id], backref='in_bundles')
    
    @property
    def profit(self):
        return self.sale_price_in_bundle - (self.product.purchase_price if self.product else 0)


ORDER_TYPE_DELIVERY = 'delivery'
ORDER_TYPE_WALKIN = 'walkin'

class Order(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=True, index=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True, index=True)
    date = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    status = db.Column(db.String(50), nullable=False, index=True)
    status_updated_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    inventory_deducted = db.Column(db.Boolean, nullable=False, default=False, index=True)
    order_type = db.Column(db.String(20), nullable=False, server_default=ORDER_TYPE_DELIVERY, index=True)

    delivery_fees = db.Column(db.Float, default=0)
    total_amount = db.Column(db.Float, default=0)
    amount_paid = db.Column(db.Float, default=0)
    cod_fee_applied = db.Column(db.Float, nullable=False, default=0)
    notes = db.Column(db.Text)
    tracking_number = db.Column(db.String(100), index=True)
    customer_called = db.Column(db.Boolean, default=False)
    customer_called_at = db.Column(db.DateTime, nullable=True)
    customer_called_by_id = db.Column(db.Integer, db.ForeignKey('employee.id', ondelete='SET NULL'), nullable=True)
    
    customer_verified = db.Column(db.Boolean, default=False)
    customer_verified_at = db.Column(db.DateTime, nullable=True)
    customer_verified_by_id = db.Column(db.Integer, db.ForeignKey('employee.id', ondelete='SET NULL'), nullable=True)
    
    registered = db.Column(db.Boolean, default=False)
    registered_at = db.Column(db.DateTime, nullable=True)
    registered_by_id = db.Column(db.Integer, db.ForeignKey('employee.id', ondelete='SET NULL'), nullable=True)
    
    weight          = db.Column(db.Float,      nullable=True)
    package_volume  = db.Column(db.String(20), nullable=True)
    delivery_notes  = db.Column(db.Text,       nullable=True)
    
    is_urgent = db.Column(db.Boolean, default=False, index=True)
    is_nearest_post_branch = db.Column(db.Boolean, default=False)
    
    # Snapshot of customer data at order creation time
    _customer_name = db.Column('customer_name', db.String(100), nullable=True)
    _customer_phone = db.Column('customer_phone', db.String(20), nullable=True)
    _customer_governorate = db.Column('customer_governorate', db.String(100), nullable=True)
    _customer_address_details = db.Column('customer_address_details', db.String(200), nullable=True)

    @property
    def customer_name(self):
        return self._customer_name or (self.customer.name if self.customer else None)

    @customer_name.setter
    def customer_name(self, value):
        self._customer_name = value

    @property
    def customer_phone(self):
        return self._customer_phone or (self.customer.phone if self.customer else None)

    @customer_phone.setter
    def customer_phone(self, value):
        self._customer_phone = value

    @property
    def customer_governorate(self):
        return self._customer_governorate or (self.customer.governorate if self.customer else None)

    @customer_governorate.setter
    def customer_governorate(self, value):
        self._customer_governorate = value

    @property
    def customer_address_details(self):
        return self._customer_address_details or (self.customer.address_details if self.customer else None)

    @customer_address_details.setter
    def customer_address_details(self, value):
        self._customer_address_details = value

    items = db.relationship('OrderItem', backref='order', lazy=True, cascade="all, delete-orphan")
    issues = db.relationship('Issue', backref='order', lazy=True, cascade="all, delete-orphan")
    logs = db.relationship('CustomerLog', backref='order', lazy=True, cascade="all, delete-orphan")
    employee = db.relationship('Employee', backref='orders', foreign_keys=[employee_id])
    
    registered_by = db.relationship('Employee', foreign_keys=[registered_by_id], backref='registered_orders')
    verified_by = db.relationship('Employee', foreign_keys=[customer_verified_by_id], backref='verified_orders')
    called_by = db.relationship('Employee', foreign_keys=[customer_called_by_id], backref='called_orders')

    @property
    def total_price(self):
        return self.total_amount

    @property
    def delivery_cost(self):
        return self.delivery_fees

    @property
    def remaining_amount(self):
        total = (self.total_amount or 0) + (self.delivery_fees or 0) + (self.cod_fee_applied or 0)
        return total - (self.amount_paid or 0)
    
    @property
    def is_walkin(self):
        return self.order_type == ORDER_TYPE_WALKIN

    @property
    def is_delivery(self):
        return self.order_type == ORDER_TYPE_DELIVERY

    def add_status_history(self, status, employee_id=None, notes=None):
        last_history = self.status_history.order_by(OrderStatusHistory.timestamp.desc()).first()
        if last_history and last_history.status == status:
            if employee_id and not last_history.changed_by_employee_id:
                last_history.changed_by_employee_id = employee_id
            if notes and not last_history.notes:
                last_history.notes = notes
            return last_history
        
        history = OrderStatusHistory(
            order_id=self.id,
            status=status,
            changed_by_employee_id=employee_id,
            notes=notes
        )
        db.session.add(history)
        return history
    
    def get_status_timeline(self):
        return self.status_history.order_by(OrderStatusHistory.timestamp.asc()).all()

    __table_args__ = (
        db.Index('ix_order_status_date', 'status', 'date'),
    )


@event.listens_for(Order, 'before_update')
def _order_update_status_timestamp(mapper, connection, target):
    try:
        insp = db.inspect(target)
        if insp.attrs.status.history.has_changes():
            target.status_updated_at = datetime.utcnow()
    except Exception:
        target.status_updated_at = datetime.utcnow()


@event.listens_for(Order, 'after_insert')
def _order_create_initial_status(mapper, connection, target):
    from sqlalchemy import text
    connection.execute(
        text("""
            INSERT INTO order_status_history (order_id, status, timestamp)
            VALUES (:order_id, :status, :timestamp)
        """),
        {"order_id": target.id, "status": target.status, "timestamp": datetime.utcnow()}
    )

class ReplacementOrder(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=False, index=True)
    original_order_id = db.Column(db.Integer, db.ForeignKey('order.id'), nullable=True, index=True)
    original_replacement_order_id = db.Column(db.Integer, db.ForeignKey('replacement_order.id'), nullable=True, index=True)
    date = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    status = db.Column(db.String(50), nullable=False, default='جديد', index=True)
    status_updated_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    inventory_deducted = db.Column(db.Boolean, nullable=False, default=False, index=True)
    delivery_fees = db.Column(db.Float, default=0)
    total_amount = db.Column(db.Float, default=0)
    delivery_fees_customer = db.Column(db.Float, default=0.0)
    amount_paid = db.Column(db.Float, default=0)
    cod_fee_applied = db.Column(db.Float, nullable=False, default=0)
    notes = db.Column(db.Text)
    tracking_number = db.Column(db.String(100))
    customer_called = db.Column(db.Boolean, default=False)
    customer_called_at = db.Column(db.DateTime, nullable=True)
    customer_called_by_id = db.Column(db.Integer, db.ForeignKey('employee.id', ondelete='SET NULL'), nullable=True)
    
    customer_verified = db.Column(db.Boolean, default=False)
    customer_verified_at = db.Column(db.DateTime, nullable=True)
    customer_verified_by_id = db.Column(db.Integer, db.ForeignKey('employee.id', ondelete='SET NULL'), nullable=True)
    
    registered = db.Column(db.Boolean, default=False)
    registered_at = db.Column(db.DateTime, nullable=True)
    registered_by_id = db.Column(db.Integer, db.ForeignKey('employee.id', ondelete='SET NULL'), nullable=True)
    
    weight          = db.Column(db.Float,      nullable=True)
    package_volume  = db.Column(db.String(20), nullable=True)
    delivery_notes  = db.Column(db.Text,       nullable=True)
    
    is_urgent = db.Column(db.Boolean, default=False, index=True)
    
    alternative_name = db.Column(db.String(200), nullable=True)
    alternative_phone = db.Column(db.String(20), nullable=True)
    alternative_governorate = db.Column(db.String(100), nullable=True)
    alternative_address_details = db.Column(db.Text, nullable=True)

    @property
    def customer_name(self):
        return self.alternative_name or (self.customer.name if self.customer else None)

    @property
    def customer_phone(self):
        return self.alternative_phone or (self.customer.phone if self.customer else None)

    @property
    def customer_governorate(self):
        return self.alternative_governorate or (self.customer.governorate if self.customer else None)

    @property
    def customer_address_details(self):
        return self.alternative_address_details or (self.customer.address_details if self.customer else None)

    items = db.relationship('ReplacementOrderItem', backref='replacement_order', lazy=True, cascade="all, delete-orphan")
    customer = db.relationship('Customer', backref='replacement_orders')
    original_replacement_order = db.relationship(
        'ReplacementOrder',
        primaryjoin='ReplacementOrder.original_replacement_order_id == remote(ReplacementOrder.id)',
        backref=db.backref('derived_replacements', lazy='dynamic')
    )
    
    registered_by = db.relationship('Employee', foreign_keys=[registered_by_id], backref='registered_replacement_orders')
    verified_by = db.relationship('Employee', foreign_keys=[customer_verified_by_id], backref='verified_replacement_orders')
    called_by = db.relationship('Employee', foreign_keys=[customer_called_by_id], backref='called_replacement_orders')

    damaged_products_loss = db.Column(db.Float, default=0.0)
    customer_refund_amount = db.Column(db.Float, default=0.0)
    delivery_fees_loss = db.Column(db.Float, default=0.0)
    total_loss = db.Column(db.Float, default=0.0)

    # Draft functionality fields
    is_draft = db.Column(db.Boolean, default=False, nullable=False, index=True)
    draft_step = db.Column(db.String(50), nullable=True)  # 'product_condition', 'new_products', 'review'
    draft_data = db.Column(db.Text, nullable=True)  # JSON string containing partial form data

    @property
    def total_price(self):
        return self.total_amount

    @property
    def delivery_cost(self):
        return self.delivery_fees

    @property
    def remaining_amount(self):
        total = (self.total_amount or 0) + (self.delivery_fees or 0) + (self.cod_fee_applied or 0)
        return total - (self.amount_paid or 0)

    @property
    def net_loss(self):
        return (self.total_loss or 0) - (self.customer_refund_amount or 0)

    def calculate_losses(self):
        total_damage_loss = 0
        total_customer_refund = 0
        
        for item in self.items:
            if item.is_damaged:
                total_damage_loss += item.purchase_price * item.quantity
                total_customer_refund += item.price * item.quantity
            elif item.state == 'مرتجع':
                total_customer_refund += item.price * item.quantity
            else:
                pass
        
        self.damaged_products_loss = total_damage_loss
        self.customer_refund_amount = total_customer_refund
        self.delivery_fees_loss = self.delivery_fees or 0
        self.total_loss = total_damage_loss + self.delivery_fees_loss

    def add_status_history(self, status, employee_id=None, notes=None):
        last_history = self.status_history.order_by(ReplacementOrderStatusHistory.timestamp.desc()).first()
        if last_history and last_history.status == status:
            if employee_id and not last_history.changed_by_employee_id:
                last_history.changed_by_employee_id = employee_id
            if notes and not last_history.notes:
                last_history.notes = notes
            return last_history
        
        history = ReplacementOrderStatusHistory(
            replacement_order_id=self.id,
            status=status,
            changed_by_employee_id=employee_id,
            notes=notes
        )
        db.session.add(history)
        return history
    
    def get_status_timeline(self):
        return self.status_history.order_by(ReplacementOrderStatusHistory.timestamp.asc()).all()

    __table_args__ = (
        db.Index('ix_replacement_order_status_date', 'status', 'date'),
    )


@event.listens_for(ReplacementOrder, 'before_update')
def _replacement_order_update_status_timestamp(mapper, connection, target):
    try:
        insp = db.inspect(target)
        if insp.attrs.status.history.has_changes():
            target.status_updated_at = datetime.utcnow()
    except Exception:
        target.status_updated_at = datetime.utcnow()


@event.listens_for(ReplacementOrder, 'after_insert')
def _replacement_order_create_initial_status(mapper, connection, target):
    from sqlalchemy import text
    connection.execute(
        text("""
            INSERT INTO replacement_order_status_history (replacement_order_id, status, timestamp)
            VALUES (:replacement_order_id, :status, :timestamp)
        """),
        {"replacement_order_id": target.id, "status": target.status, "timestamp": datetime.utcnow()}
    )


class ReplacementOrderItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    replacement_order_id = db.Column(db.Integer, db.ForeignKey('replacement_order.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id', ondelete='SET NULL'), nullable=True)
    variant_id = db.Column(db.Integer, db.ForeignKey('product_variant.id', ondelete='SET NULL', name='fk_replacementorderitem_variant_id'), nullable=True)
    
    size_variant_id = db.Column(db.Integer, db.ForeignKey('product_variant.id', ondelete='SET NULL', name='fk_replacementorderitem_size_variant_id'), nullable=True)
    color_variant_id = db.Column(db.Integer, db.ForeignKey('product_variant.id', ondelete='SET NULL', name='fk_replacementorderitem_color_variant_id'), nullable=True)
    style_variant_id = db.Column(db.Integer, db.ForeignKey('product_variant.id', ondelete='SET NULL', name='fk_replacementorderitem_style_variant_id'), nullable=True)
    
    quantity = db.Column(db.Integer, default=1)
    price = db.Column(db.Float, nullable=False)
    state = db.Column(db.String(20))
    
    is_damaged = db.Column(db.Boolean, default=False)
    is_returned = db.Column(db.Boolean, default=False)
    purchase_price = db.Column(db.Float, default=0.0)
    damage_loss = db.Column(db.Float, default=0.0)

    _product = db.relationship('Product', backref='replacement_order_items', foreign_keys=[product_id])
    variant = db.relationship('ProductVariant', backref='replacement_order_items', foreign_keys=[variant_id])
    size_variant = db.relationship('ProductVariant', backref='size_replacement_order_items', foreign_keys=[size_variant_id])
    color_variant = db.relationship('ProductVariant', backref='color_replacement_order_items', foreign_keys=[color_variant_id])
    style_variant = db.relationship('ProductVariant', backref='style_replacement_order_items', foreign_keys=[style_variant_id])
    
    @property
    def product(self):
        """
        Resolve the product for this item.

        WARNING: This property may trigger additional DB queries when called on
        items loaded without eager loading. Callers iterating over many items
        should use joinedload() or selectinload() on the relevant relationships
        to avoid N+1 query patterns.

        Returns:
            Product instance or None if the product no longer exists.
        """
        if self._product:
            return self._product
        
        if self.size_variant and self.size_variant.product:
            return self.size_variant.product
        elif self.color_variant and self.color_variant.product:
            return self.color_variant.product
        elif self.style_variant and self.style_variant.product:
            return self.style_variant.product
        elif self.variant and self.variant.product:
            return self.variant.product
            
        return None
    
    @property
    def selected_variants(self):
        variants = {}
        
        if self.size_variant:
            variants['مقاس'] = self.size_variant
        if self.color_variant:
            variants['لون'] = self.color_variant
        if self.style_variant:
            variants['شكل'] = self.style_variant
            
        if not variants and self.variant:
            if self.variant.group_name == 'مقاس':
                variants['مقاس'] = self.variant
            elif self.variant.group_name == 'لون':
                variants['لون'] = self.variant
            elif self.variant.group_name == 'شكل':
                variants['شكل'] = self.variant
            else:
                variants[self.variant.group_name] = self.variant
                
        return variants
    
    @property
    def bundle_variants_map(self) -> dict:
        return _parse_bundle_variants_map(getattr(self, 'bundle_variants_json', None))
    
    @property
    def is_valid_product(self):
        prod = self.product
        if prod is None:
            return False
        return not prod.is_deleted

    @property
    def calculated_damage_loss(self):
        if self.is_damaged:
            return (self.purchase_price or 0) * self.quantity
        return 0

    @property
    def customer_refund_value(self):
        if not self.is_damaged:
            return (self.price or 0) * self.quantity
        return 0


class OrderStatusHistory(db.Model):
    __tablename__ = 'order_status_history'
    
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('order.id', ondelete='CASCADE'), nullable=False, index=True)
    status = db.Column(db.String(50), nullable=False, index=True)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)
    changed_by_employee_id = db.Column(db.Integer, db.ForeignKey('employee.id', ondelete='SET NULL'), nullable=True)
    notes = db.Column(db.Text, nullable=True)
    
    order = db.relationship('Order', backref=backref('status_history', lazy='dynamic', cascade='all, delete-orphan', order_by='OrderStatusHistory.timestamp'))
    changed_by = db.relationship('Employee', foreign_keys=[changed_by_employee_id])
    
    __table_args__ = (
        db.Index('ix_order_status_history_order_timestamp', 'order_id', 'timestamp'),
        db.Index('ix_order_status_history_status_timestamp', 'status', 'timestamp'),
    )
    
    def __repr__(self):
        return f'<OrderStatusHistory {self.order_id}: {self.status} at {self.timestamp}>'


class ReplacementOrderStatusHistory(db.Model):
    __tablename__ = 'replacement_order_status_history'
    
    id = db.Column(db.Integer, primary_key=True)
    replacement_order_id = db.Column(db.Integer, db.ForeignKey('replacement_order.id', ondelete='CASCADE'), nullable=False, index=True)
    status = db.Column(db.String(50), nullable=False, index=True)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)
    changed_by_employee_id = db.Column(db.Integer, db.ForeignKey('employee.id', ondelete='SET NULL'), nullable=True)
    notes = db.Column(db.Text, nullable=True)
    
    replacement_order = db.relationship('ReplacementOrder', backref=backref('status_history', lazy='dynamic', cascade='all, delete-orphan', order_by='ReplacementOrderStatusHistory.timestamp'))
    changed_by = db.relationship('Employee', foreign_keys=[changed_by_employee_id])
    
    __table_args__ = (
        db.Index('ix_replacement_order_status_history_order_timestamp', 'replacement_order_id', 'timestamp'),
        db.Index('ix_replacement_order_status_history_status_timestamp', 'status', 'timestamp'),
    )
    
    def __repr__(self):
        return f'<ReplacementOrderStatusHistory {self.replacement_order_id}: {self.status} at {self.timestamp}>'


class ReturnOrder(db.Model):
    __tablename__ = 'return_order'
    id = db.Column(db.Integer, primary_key=True)
    original_order_id = db.Column(db.Integer, db.ForeignKey('order.id'), nullable=True)
    original_replacement_order_id = db.Column(db.Integer, db.ForeignKey('replacement_order.id'), nullable=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=False, index=True)
    customer_refund_amount = db.Column(db.Float, nullable=False, default=0.0)
    received_by_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    __table_args__ = (
        db.CheckConstraint(
            '(original_order_id IS NOT NULL) + '
            '(original_replacement_order_id IS NOT NULL) = 1',
            name='ck_return_order_single_source'
        ),
        db.Index('ix_return_order_customer_created', 'customer_id', 'created_at'),
        db.Index('ix_return_order_source_order', 'original_order_id'),
        db.Index('ix_return_order_source_replacement', 'original_replacement_order_id'),
    )

    items = db.relationship('ReturnOrderItem', backref='return_order', lazy='dynamic',
                            cascade='all, delete-orphan')
    received_by = db.relationship('Employee', foreign_keys=[received_by_id])
    customer = db.relationship('Customer', backref=db.backref('return_orders', lazy='dynamic'))
    original_order = db.relationship('Order', foreign_keys=[original_order_id],
                                     backref=db.backref('return_orders', lazy='dynamic'))
    original_replacement_order = db.relationship('ReplacementOrder', foreign_keys=[original_replacement_order_id],
                                                  backref=db.backref('return_orders', lazy='dynamic'))


class ReturnOrderItem(db.Model):
    __tablename__ = 'return_order_item'
    id = db.Column(db.Integer, primary_key=True)
    return_order_id = db.Column(db.Integer, db.ForeignKey('return_order.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=False, index=True)
    quantity = db.Column(db.Integer, nullable=False, default=1)
    item_condition = db.Column(db.String(20), nullable=False)
    return_reason = db.Column(db.String(50), nullable=False)
    unit_sale_price_snapshot = db.Column(db.Float, nullable=False)
    unit_purchase_price_snapshot = db.Column(db.Float, nullable=False)
    inspected_by_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    inspected_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    __table_args__ = (
        db.Index('ix_return_order_item_product', 'product_id', 'return_order_id'),
    )

    product = db.relationship('Product', foreign_keys=[product_id])
    inspected_by = db.relationship('Employee', foreign_keys=[inspected_by_id])


class OrderItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('order.id'), nullable=False, index=True)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id', ondelete='SET NULL'), nullable=True)
    variant_id = db.Column(db.Integer, db.ForeignKey('product_variant.id', ondelete='SET NULL', name='fk_orderitem_variant_id'), nullable=True)
    
    size_variant_id = db.Column(db.Integer, db.ForeignKey('product_variant.id', ondelete='SET NULL', name='fk_orderitem_size_variant_id'), nullable=True)
    color_variant_id = db.Column(db.Integer, db.ForeignKey('product_variant.id', ondelete='SET NULL', name='fk_orderitem_color_variant_id'), nullable=True)
    style_variant_id = db.Column(db.Integer, db.ForeignKey('product_variant.id', ondelete='SET NULL', name='fk_orderitem_style_variant_id'), nullable=True)
    
    quantity = db.Column(db.Integer, default=1)
    price = db.Column(db.Float, nullable=False)
    purchase_price_snapshot = db.Column(db.Float, nullable=True)
    state = db.Column(db.String(50), default='لم يتم التجربة بعد')
    
    bundle_variants_json = db.Column(db.Text, nullable=True)

    _product = db.relationship('Product', backref='order_items', foreign_keys=[product_id])
    variant = db.relationship('ProductVariant', backref='order_items', foreign_keys=[variant_id])
    size_variant = db.relationship('ProductVariant', backref='size_order_items', foreign_keys=[size_variant_id])
    color_variant = db.relationship('ProductVariant', backref='color_order_items', foreign_keys=[color_variant_id])
    style_variant = db.relationship('ProductVariant', backref='style_order_items', foreign_keys=[style_variant_id])
    
    @property
    def selected_variants(self):
        variants = {}
        
        if self.size_variant:
            variants['مقاس'] = self.size_variant
        if self.color_variant:
            variants['لون'] = self.color_variant
        if self.style_variant:
            variants['شكل'] = self.style_variant
            
        if not variants and self.variant:
            if self.variant.group_name == 'مقاس':
                variants['مقاس'] = self.variant
            elif self.variant.group_name == 'لون':
                variants['لون'] = self.variant
            elif self.variant.group_name == 'شكل':
                variants['شكل'] = self.variant
            else:
                variants[self.variant.group_name] = self.variant
                
        return variants
    
    @property
    def product(self):
        """
        Resolve the product for this item.

        WARNING: This property may trigger additional DB queries when called on
        items loaded without eager loading. Callers iterating over many items
        should use joinedload() or selectinload() on the relevant relationships
        to avoid N+1 query patterns.

        Returns:
            Product instance or None if the product no longer exists.
        """
        if self._product:
            return self._product
        
        if self.size_variant and self.size_variant.product:
            return self.size_variant.product
        elif self.color_variant and self.color_variant.product:
            return self.color_variant.product
        elif self.style_variant and self.style_variant.product:
            return self.style_variant.product
        
        elif self.variant and self.variant.product:
            return self.variant.product
        
        # Fallback DB query in case all relationship attributes are unloaded.
        # ⚠️ N+1 RISK: this fires a query per item if relationships are not eager-loaded.
        try:
            from . import db
            variant = db.session.query(ProductVariant).filter(
                (ProductVariant.id == self.variant_id) |
                (ProductVariant.id == self.size_variant_id) |
                (ProductVariant.id == self.color_variant_id) |
                (ProductVariant.id == self.style_variant_id)
            ).first()
            
            if variant and variant.product:
                return variant.product
        except Exception:
            # Silently ignore DB errors in the variant fallback path.
            # ⚠️ If this fires frequently, check for missing relationships or deleted variants.
            pass

        return None

    @property
    def bundle_variants_map(self) -> dict:
        return _parse_bundle_variants_map(self.bundle_variants_json)
    
    @property
    def is_valid_product(self):
        product = self.product
        if product is None:
            return False
        return not product.is_deleted


class Issue(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=False, index=True)
    order_id = db.Column(db.Integer, db.ForeignKey('order.id'), nullable=False, index=True)
    new_tracking_number = db.Column(db.String(100), nullable=False)
    state = db.Column(db.String(50), default='قيد المعالجة', index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    notes = db.Column(db.Text, nullable=True)
    
    customer = db.relationship('Customer', backref='issues')
    items = db.relationship('IssueItem', backref='issue', lazy=True, cascade="all, delete-orphan")

class IssueItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    issue_id = db.Column(db.Integer, db.ForeignKey('issue.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=True)
    product_name = db.Column(db.String(100), nullable=False)
    product_state = db.Column(db.String(50), default='لم يتم التجربة بعد')
    price = db.Column(db.Float, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    

class FollowUp(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id', ondelete='CASCADE', name='fk_followup_customer_id'), nullable=False)
    problem = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(20), nullable=False, default='قائمة')
    created_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    last_contact_at = db.Column(db.DateTime, nullable=True)
    next_contact_due = db.Column(db.DateTime, nullable=True, index=True)

    def __repr__(self):
        return f'<FollowUp id={self.id} customer_id={self.customer_id} status={self.status}>'

    @property
    def is_open(self):
        return self.status == 'قائمة'


class CustomerLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    customer_id = db.Column(
        db.Integer,
        db.ForeignKey('customer.id', ondelete="CASCADE", name='fk_customer_log_customer_id'),
        nullable=False
    )

    order_id = db.Column(
        db.Integer,
        db.ForeignKey('order.id', name='fk_customer_log_order_id'),
        nullable=True
    )

    return_order_id = db.Column(
        db.Integer,
        db.ForeignKey('return_order.id'),
        nullable=True,
        index=True
    )

    employee_id = db.Column(
        db.Integer,
        db.ForeignKey('employee.id', ondelete='SET NULL', name='fk_customer_log_employee_id'),
        nullable=True,
        index=True
    )

    timestamp = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    type = db.Column(db.String(50), nullable=False, index=True)
    content = db.Column(db.Text, nullable=False)
    
    reminder_time = db.Column(db.DateTime, nullable=True)
    reminder_duration = db.Column(db.Integer, nullable=True)
    reminder_duration_type = db.Column(db.String(20), nullable=True)
    follow_up_reason = db.Column(db.Text, nullable=True)
    is_dismissed = db.Column(db.Boolean, default=False)

    employee = db.relationship('Employee', foreign_keys=[employee_id])


class OrderEditLog(db.Model):
    """Order edit log that records each save with the responsible employee."""
    __tablename__ = 'order_edit_log'

    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('order.id', ondelete='CASCADE'), nullable=True, index=True)
    replacement_order_id = db.Column(db.Integer, db.ForeignKey('replacement_order.id', ondelete='CASCADE'), nullable=True, index=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id', ondelete='SET NULL'), nullable=True)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)
    description = db.Column(db.Text, nullable=False)

    employee = db.relationship('Employee', foreign_keys=[employee_id])
    order = db.relationship('Order', foreign_keys=[order_id],
                            backref=backref('edit_logs', lazy='dynamic',
                                            cascade='all, delete-orphan',
                                            order_by='OrderEditLog.timestamp'))
    replacement_order = db.relationship('ReplacementOrder', foreign_keys=[replacement_order_id],
                                        backref=backref('edit_logs', lazy='dynamic',
                                                        cascade='all, delete-orphan',
                                                        order_by='OrderEditLog.timestamp'))

    def __repr__(self):
        return f'<OrderEditLog order_id={self.order_id} emp={self.employee_id} at {self.timestamp}>'


class GovernorateFee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    fee = db.Column(db.Float, nullable=False, default=0)


class AppSettings(db.Model):
    __tablename__ = 'app_settings'
    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(100), unique=True, nullable=False)
    value = db.Column(db.Text, nullable=True)
    description = db.Column(db.String(255), nullable=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    @classmethod
    def get_value(cls, key, default=None):
        setting = cls.query.filter_by(key=key).first()
        if setting:
            return setting.value
        return default
    
    @classmethod
    def set_value(cls, key, value, description=None):
        setting = cls.query.filter_by(key=key).first()
        if setting:
            setting.value = str(value)
            setting.updated_at = datetime.utcnow()
        else:
            setting = cls(key=key, value=str(value), description=description)
            db.session.add(setting)
        return setting
    
    @classmethod
    def get_cod_fee(cls):
        value = cls.get_value('cod_fee', '0')
        try:
            return float(value)
        except (ValueError, TypeError):
            return 0.0


class Employee(db.Model):
    national_id = db.Column(db.String(14), nullable=True)
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    phone = db.Column(db.String(20), nullable=False, unique=True)
    email = db.Column(db.String(100), nullable=True)
    address = db.Column(db.String(200), nullable=True)
    position = db.Column(db.String(100), nullable=False)
    hire_date = db.Column(db.Date, nullable=False, default=date.today)
    salary = db.Column(db.Float, nullable=False, default=0)
    sales_commission_percentage = db.Column(db.Float, nullable=False, default=0)
    is_active = db.Column(db.Boolean, default=True, index=True)
    username = db.Column(db.String(50), nullable=False, unique=True)
    password_hash = db.Column(db.String(255), nullable=False)
    
    can_view_customers = db.Column(db.Boolean, default=False)
    can_add_customers = db.Column(db.Boolean, default=False)
    can_edit_customers = db.Column(db.Boolean, default=False)
    can_delete_customers = db.Column(db.Boolean, default=False)
    can_view_customer_orders = db.Column(db.Boolean, default=False)
    can_view_customer_logs = db.Column(db.Boolean, default=False)
    can_view_customer_profile = db.Column(db.Boolean, default=False)
    can_add_customer_logs = db.Column(db.Boolean, default=False)
    can_delete_customer_logs = db.Column(db.Boolean, default=False)
    
    can_view_products = db.Column(db.Boolean, default=False)
    can_add_products = db.Column(db.Boolean, default=False)
    can_edit_products = db.Column(db.Boolean, default=False)
    can_delete_products = db.Column(db.Boolean, default=False)
    can_restore_products = db.Column(db.Boolean, default=False)
    can_delete_all_products = db.Column(db.Boolean, default=False)
    can_view_purchase_price = db.Column(db.Boolean, default=False)
    can_view_sold_products_by_quantity = db.Column(db.Boolean, default=False)
    
    can_view_orders = db.Column(db.Boolean, default=False)
    can_add_orders = db.Column(db.Boolean, default=False)
    can_edit_orders = db.Column(db.Boolean, default=False)
    can_delete_orders = db.Column(db.Boolean, default=False)
    can_update_order_status = db.Column(db.Boolean, default=False)
    can_update_tracking = db.Column(db.Boolean, default=False)
    can_view_orders_by_status = db.Column(db.Boolean, default=False)
    
    can_view_replacements = db.Column(db.Boolean, default=False)
    can_add_replacements = db.Column(db.Boolean, default=False)
    can_edit_replacements = db.Column(db.Boolean, default=False)
    can_delete_replacements = db.Column(db.Boolean, default=False)
    can_view_replacements_by_state = db.Column(db.Boolean, default=False)

    can_view_returns = db.Column(db.Boolean, default=False)
    can_add_returns = db.Column(db.Boolean, default=False)
    can_view_returns_by_state = db.Column(db.Boolean, default=False)
    can_delete_returns = db.Column(db.Boolean, default=False)
    
    can_view_fees = db.Column(db.Boolean, default=False)
    can_edit_fees = db.Column(db.Boolean, default=False)
    can_delete_fees = db.Column(db.Boolean, default=False)
    
    can_view_employees = db.Column(db.Boolean, default=False)
    can_add_employees = db.Column(db.Boolean, default=False)
    can_edit_employees = db.Column(db.Boolean, default=False)
    can_delete_employees = db.Column(db.Boolean, default=False)
    can_view_employee_logs = db.Column(db.Boolean, default=False)
    can_manage_employee_salary = db.Column(db.Boolean, default=False)
    can_delete_salary_transactions = db.Column(db.Boolean, default=False)
    can_view_employee_activity = db.Column(db.Boolean, default=False)
    
    can_view_statistics = db.Column(db.Boolean, default=False)
    # Statistics cards
    can_view_stats_stock = db.Column(db.Boolean, default=False)
    can_view_stats_fixed_assets = db.Column(db.Boolean, default=False)
    can_view_stats_total_debt = db.Column(db.Boolean, default=False)
    can_view_stats_capital_growth = db.Column(db.Boolean, default=False)
    can_view_stats_daily = db.Column(db.Boolean, default=False)
    can_view_stats_pending_orders = db.Column(db.Boolean, default=False)
    can_view_stats_net_profit = db.Column(db.Boolean, default=False)
    can_view_stats_losses = db.Column(db.Boolean, default=False)
    can_view_stats_delivery_rate = db.Column(db.Boolean, default=False)
    can_view_stats_sales = db.Column(db.Boolean, default=False)
    can_view_stats_amount_paid = db.Column(db.Boolean, default=False)
    can_view_stats_monthly_delivered = db.Column(db.Boolean, default=False)
    can_view_stats_monthly_invoices = db.Column(db.Boolean, default=False)
    can_view_stats_fixed_assets_expenses = db.Column(db.Boolean, default=False)
    can_view_stats_employee_salaries = db.Column(db.Boolean, default=False)
    can_view_stats_employee_debt = db.Column(db.Boolean, default=False)
    can_view_stats_operational_expenses = db.Column(db.Boolean, default=False)
    
    can_view_suppliers = db.Column(db.Boolean, default=False)
    can_add_suppliers = db.Column(db.Boolean, default=False)
    can_edit_suppliers = db.Column(db.Boolean, default=False)
    can_delete_suppliers = db.Column(db.Boolean, default=False)
    can_view_supplier_invoices = db.Column(db.Boolean, default=False)
    can_view_supplier_history = db.Column(db.Boolean, default=False)
    can_pay_supplier_debt = db.Column(db.Boolean, default=False)
    can_add_supplier_debt = db.Column(db.Boolean, default=False)
    
    can_view_invoices = db.Column(db.Boolean, default=False)
    can_add_invoices = db.Column(db.Boolean, default=False)
    can_edit_invoices = db.Column(db.Boolean, default=False)
    can_delete_invoices = db.Column(db.Boolean, default=False)
    can_view_invoice_details = db.Column(db.Boolean, default=False)
    
    can_view_expenses = db.Column(db.Boolean, default=False)
    can_add_expenses = db.Column(db.Boolean, default=False)
    can_edit_expenses = db.Column(db.Boolean, default=False)
    can_delete_expenses = db.Column(db.Boolean, default=False)
    can_view_operational_expenses = db.Column(db.Boolean, default=False)
    can_view_fixed_assets = db.Column(db.Boolean, default=False)
    
    can_view_activity_log = db.Column(db.Boolean, default=False)
    can_manage_reminders = db.Column(db.Boolean, default=False)
    can_manage_stocktake = db.Column(db.Boolean, default=False)
    is_admin = db.Column(db.Boolean, default=False)

    requires_attendance = db.Column(db.Boolean, default=False, server_default='false', nullable=False)
    can_manage_attendance = db.Column(db.Boolean, default=False, server_default='false', nullable=False)

    can_view_followups = db.Column(db.Boolean, default=False)
    can_add_followups = db.Column(db.Boolean, default=False)
    can_edit_followups = db.Column(db.Boolean, default=False)
    can_delete_followups = db.Column(db.Boolean, default=False)

    can_view_damaged_products = db.Column(db.Boolean, default=False)
    can_add_damaged_products = db.Column(db.Boolean, default=False)
    can_delete_damaged_products = db.Column(db.Boolean, default=False)
    
    can_view_transactions = db.Column(db.Boolean, default=False)
    can_add_transactions = db.Column(db.Boolean, default=False)
    can_edit_transactions = db.Column(db.Boolean, default=False)
    can_delete_transactions = db.Column(db.Boolean, default=False)
    
    salary_transactions = db.relationship('SalaryTransaction', backref='employee', lazy=True, cascade="all, delete-orphan", foreign_keys='SalaryTransaction.employee_id')
    login_logs = db.relationship('EmployeeLoginLog', backref='employee', lazy=True, cascade="all, delete-orphan")
    activity_logs = db.relationship('EmployeeActivityLog', backref='employee', lazy=True, cascade="all, delete-orphan")

    def set_password(self, password):
        from werkzeug.security import generate_password_hash
        self.password_hash = generate_password_hash(password)
    
    def check_password(self, password):
        from werkzeug.security import check_password_hash
        return check_password_hash(self.password_hash, password)

    def _query_period_orders(self, start_date, end_date) -> list:
        """
        Fetch orders for this employee that were marked as 'وصل' within the given
        date range, using OrderStatusHistory as the source of truth.

        This correctly handles orders whose status later changed (e.g. to 'استبدال'):
        they are still counted in the month they were originally delivered.

        Args:
            start_date: datetime — inclusive start of the period.
            end_date: datetime — exclusive end of the period.

        Returns:
            List of Order objects that reached 'وصل' status within the period.
        """
        first_delivered_times = db.session.query(
            OrderStatusHistory.order_id,
            db.func.min(OrderStatusHistory.timestamp).label('first_delivered_at')
        ).filter(
            OrderStatusHistory.status == 'وصل'
        ).group_by(
            OrderStatusHistory.order_id
        ).subquery()

        delivered_order_ids = db.session.query(first_delivered_times.c.order_id).filter(
            first_delivered_times.c.first_delivered_at >= start_date,
            first_delivered_times.c.first_delivered_at < end_date
        ).subquery()

        return Order.query.filter(
            Order.employee_id == self.id,
            Order.status.in_(['وصل', 'استبدال']),
            Order.id.in_(delivered_order_ids)
        ).all()

    def _calculate_sales_commission(self, month: int | None, year: int | None) -> dict:
        """
        Calculate sales commission for the given period.

        Args:
            month: 1-12 or None for all-time.
            year: Full year int or None for all-time.

        Returns:
            Dict containing commission details:
                commission: Commission amount as float.
                total_sales: Sum of order total_amount values for the period.
                order_count: Number of delivered/replacement orders in the period.
            Returns zeros when no commission rate is set or no period is specified.
        """
        if self.sales_commission_percentage <= 0 or not (month and year):
            return {
                'commission': 0.0,
                'total_sales': 0.0,
                'order_count': 0,
            }

        start_date = datetime(year, month, 1)
        end_date = datetime(year + 1, 1, 1) if month == 12 else datetime(year, month + 1, 1)
        period_orders = self._query_period_orders(start_date, end_date)
        total_sales = sum((order.total_amount or 0) for order in period_orders)

        return {
            'commission': (total_sales * self.sales_commission_percentage) / 100,
            'total_sales': float(total_sales),
            'order_count': len(period_orders),
        }

    @staticmethod
    def _calc_accrued_salary_attendance(net_salary, employee_id, month, year, days_in_month):
        today = date.today()
        start = date(year, month, 1)
        if (year, month) == (today.year, today.month):
            end = today
        else:
            end = date(year, month, days_in_month)

        present = AttendanceRecord.query.filter(
            AttendanceRecord.employee_id == employee_id,
            AttendanceRecord.status == 'present',
            AttendanceRecord.date >= start,
            AttendanceRecord.date <= end
        ).count()
        half = AttendanceRecord.query.filter(
            AttendanceRecord.employee_id == employee_id,
            AttendanceRecord.status == 'half_day',
            AttendanceRecord.date >= start,
            AttendanceRecord.date <= end
        ).count()

        day_rate = net_salary / days_in_month
        return day_rate * (present + 0.5 * half)

    @staticmethod
    def _calc_accrued_salary(net_salary, month, year, days_in_month):
        today = datetime.now()
        selected = (year, month)
        current = (today.year, today.month)
        if selected == current:
            accrued_days = today.day
        elif selected < current:
            accrued_days = days_in_month
        else:
            accrued_days = 0
        return (net_salary / days_in_month * accrued_days) if days_in_month > 0 and accrued_days > 0 else 0
    
    def get_salary_summary(self, month=None, year=None) -> dict:
        """
        Build a complete salary summary for the employee for the given period.

        Args:
            month: 1-12 or None for all-time totals.
            year: Full year int or None for all-time totals.

        Returns:
            Dict containing: base_salary, sales_commission, total_salary,
            total_bonuses, total_deductions, total_advances, net_salary,
            transactions, delivered_sales, delivered_orders_count.
        """
        query = SalaryTransaction.query.filter_by(employee_id=self.id)
        if month and year:
            query = query.filter_by(month=month, year=year)

        transactions = query.all()
        commission_data = self._calculate_sales_commission(month, year)
        base_salary = self.salary
        total_bonuses = sum(t.amount for t in transactions if t.transaction_type == 'مكافأة')
        total_deductions = sum(t.amount for t in transactions if t.transaction_type == 'خصم من المرتب')
        total_advances = sum(t.amount for t in transactions if t.transaction_type == 'سلفه')
        total_salary = base_salary + commission_data['commission']
        net_salary = total_salary + total_bonuses - total_deductions - total_advances
        return {'base_salary': base_salary, 'sales_commission': commission_data['commission'], 'total_salary': total_salary, 'total_bonuses': total_bonuses, 'total_deductions': total_deductions, 'total_advances': total_advances, 'net_salary': net_salary, 'transactions': transactions, 'delivered_sales': commission_data['total_sales'], 'delivered_orders_count': commission_data['order_count']}
    
    def get_monthly_transactions(self, month, year):
        return SalaryTransaction.query.filter_by(
            employee_id=self.id,
            month=month,
            year=year
        ).order_by(SalaryTransaction.transaction_date.desc()).all()

    def get_running_balance(self, since_date=None):
        since_date = since_date or ATTENDANCE_FEATURE_START_DATE
        total_accrued = 0
        cursor = date(since_date.year, since_date.month, 1)
        today = date.today()
        # ponytail: exclude current month — "previous months" means up to last month
        while (cursor.year, cursor.month) < (today.year, today.month):
            days_in_month = monthrange(cursor.year, cursor.month)[1]
            net_salary = self.get_salary_summary(cursor.month, cursor.year)['net_salary']
            if self.requires_attendance:
                accrued = Employee._calc_accrued_salary_attendance(
                    net_salary, self.id, cursor.month, cursor.year, days_in_month)
            else:
                accrued = Employee._calc_accrued_salary(
                    net_salary, cursor.month, cursor.year, days_in_month)
            total_accrued += accrued
            cursor = date(cursor.year + (cursor.month == 12), cursor.month % 12 + 1, 1)

        total_paid = db.session.query(db.func.sum(SalaryTransaction.amount)).filter(
            SalaryTransaction.employee_id == self.id,
            SalaryTransaction.transaction_type.in_(('دفع مستحق من الشهور السابقة',)),
        ).scalar() or 0

        carry_forward = db.session.query(db.func.sum(SalaryTransaction.amount)).filter(
            SalaryTransaction.employee_id == self.id,
            SalaryTransaction.transaction_type == 'باقي مستحق من الشهور السابقة',
        ).scalar() or 0

        return total_accrued + carry_forward - total_paid

    def get_previous_months_dues(self):
        carry_forward = db.session.query(db.func.sum(SalaryTransaction.amount)).filter(
            SalaryTransaction.employee_id == self.id,
            SalaryTransaction.transaction_type == 'باقي مستحق من الشهور السابقة',
        ).scalar() or 0
        paid = db.session.query(db.func.sum(SalaryTransaction.amount)).filter(
            SalaryTransaction.employee_id == self.id,
            SalaryTransaction.transaction_type == 'دفع مستحق من الشهور السابقة',
        ).scalar() or 0
        return round(carry_forward - paid, 2)


ATTENDANCE_FEATURE_START_DATE = date(2026, 7, 1)

class SalaryTransaction(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    transaction_date = db.Column(db.Date, nullable=False, default=date.today)
    transaction_type = db.Column(db.String(50), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    description = db.Column(db.Text, nullable=True)
    month = db.Column(db.Integer, nullable=True)
    year = db.Column(db.Integer, nullable=True)
    created_by = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    created_by_employee = db.relationship('Employee', foreign_keys=[created_by], backref='created_salary_transactions')
    
    def __repr__(self):
        return f'<SalaryTransaction {self.transaction_type}: {self.amount}>'
    
    @property
    def is_positive(self):
        return self.transaction_type in ('مكافأة', 'باقي مستحق من الشهور السابقة')
    
    @property
    def formatted_amount(self):
        sign = "+" if self.is_positive else "-"
        return f"{sign}{self.amount:.2f}"


class AttendanceRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    date = db.Column(db.Date, nullable=False, default=date.today)
    status = db.Column(db.String(20), nullable=False)  # present / absent / half_day
    created_by = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    __table_args__ = (
        db.UniqueConstraint('employee_id', 'date', name='uq_attendance_employee_date'),
    )

    employee = db.relationship('Employee', foreign_keys=[employee_id], backref='attendance_records')
    recorder = db.relationship('Employee', foreign_keys=[created_by])

    def __repr__(self):
        return f'<AttendanceRecord emp={self.employee_id} date={self.date} status={self.status}>'


class EmployeeLoginLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    login_time = db.Column(db.DateTime, default=datetime.utcnow)
    logout_time = db.Column(db.DateTime, nullable=True)
    ip_address = db.Column(db.String(45), nullable=True)
    user_agent = db.Column(db.String(500), nullable=True)
    

class EmployeeActivityLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    action = db.Column(db.String(100), nullable=False)
    entity_type = db.Column(db.String(50), nullable=False)
    entity_id = db.Column(db.Integer, nullable=True)
    entity_name = db.Column(db.String(200), nullable=True)
    details = db.Column(db.Text, nullable=True)
    ip_address = db.Column(db.String(45), nullable=True)
    user_agent = db.Column(db.String(500), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    
    def __repr__(self):
        return f'<EmployeeActivityLog {self.employee.name} - {self.action} {self.entity_type}>'


class Supplier(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    phone = db.Column(db.String(20), nullable=False)
    address = db.Column(db.String(200), nullable=True)
    email = db.Column(db.String(100), nullable=True)
    total_debt = db.Column(db.Float, default=0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    invoices = db.relationship('Invoice', backref='supplier', lazy=True, cascade="all, delete-orphan")


class Invoice(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    supplier_id = db.Column(db.Integer, db.ForeignKey('supplier.id'), nullable=False, index=True)
    invoice_number = db.Column(db.String(50), nullable=False, index=True)
    date = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    total_amount = db.Column(db.Float, nullable=False)
    paid_amount = db.Column(db.Float, default=0)
    notes = db.Column(db.Text, nullable=True)
    submission_id = db.Column(db.String(36), nullable=False, unique=True, index=True)
    
    items = db.relationship('InvoiceItem', backref='invoice', lazy=True, cascade="all, delete-orphan")


class InvoiceItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    invoice_id = db.Column(db.Integer, db.ForeignKey('invoice.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    purchase_price = db.Column(db.Float, nullable=False)
    selling_price = db.Column(db.Float, nullable=False)
    

class SupplierReturn(db.Model):
    __tablename__ = 'supplier_returns'
    
    id = db.Column(db.Integer, primary_key=True)
    invoice_id = db.Column(db.Integer, db.ForeignKey('invoice.id'), nullable=False, index=True)
    return_date = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    total_amount = db.Column(db.Float, default=0.0)
    notes = db.Column(db.Text, nullable=True)
    created_by = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)
    
    invoice = db.relationship('Invoice', backref='returns')
    items = db.relationship('SupplierReturnItem', backref='return_record', lazy=True, cascade='all, delete-orphan')
    created_by_employee = db.relationship('Employee', foreign_keys=[created_by], backref='supplier_returns')
    
    def __repr__(self):
        return f'<SupplierReturn invoice_id={self.invoice_id} total={self.total_amount}>'


class SupplierDebt(db.Model):
    __tablename__ = 'supplier_debts'
    
    id = db.Column(db.Integer, primary_key=True)
    supplier_id = db.Column(db.Integer, db.ForeignKey('supplier.id'), nullable=False, index=True)
    amount = db.Column(db.Float, nullable=False)
    paid_amount = db.Column(db.Float, default=0)
    date = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    notes = db.Column(db.Text, nullable=True)
    is_payment = db.Column(db.Boolean, default=False, index=True)
    created_by = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)
    
    supplier = db.relationship('Supplier', backref=backref('debts', lazy=True, cascade="all, delete-orphan"))
    created_by_employee = db.relationship('Employee', foreign_keys=[created_by], backref='created_supplier_debts')
    
    def __repr__(self):
        return f'<SupplierDebt supplier_id={self.supplier_id} amount={self.amount}>'


class SupplierReturnItem(db.Model):
    __tablename__ = 'supplier_return_items'
    
    id = db.Column(db.Integer, primary_key=True)
    return_id = db.Column(db.Integer, db.ForeignKey('supplier_returns.id'), nullable=False, index=True)
    invoice_item_id = db.Column(db.Integer, db.ForeignKey('invoice_item.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    price = db.Column(db.Float, nullable=False)
    total = db.Column(db.Float, nullable=False)
    
    invoice_item = db.relationship('InvoiceItem')
    product = db.relationship('Product')
    
    def __repr__(self):
        return f'<SupplierReturnItem product_id={self.product_id} qty={self.quantity}>'


class Expense(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    category = db.Column(db.String(100), nullable=False, index=True)
    subcategory = db.Column(db.String(100), nullable=True)
    date = db.Column(db.Date, nullable=False, default=date.today, index=True)
    description = db.Column(db.Text, nullable=True)
    receipt_number = db.Column(db.String(100), nullable=True)
    supplier = db.Column(db.String(100), nullable=True)
    payment_method = db.Column(db.String(50), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    created_by = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)
    
    created_by_employee = db.relationship('Employee', backref='created_expenses')
    
    def __repr__(self):
        return f'<Expense {self.title}: {self.amount} ج.م>'
    
    @property
    def formatted_amount(self):
        return f"{self.amount:,.2f} ج.م"
    
    @property
    def formatted_date(self):
        return self.date.strftime("%Y-%m-%d")
    
    def get_related_expenses(self):
        if self.category == 'أصول ثابتة' and self.payment_method and 'مقسوم على' in self.payment_method:
            related_expenses = Expense.query.filter(
                Expense.title == self.title,
                Expense.category == self.category,
                Expense.payment_method == self.payment_method,
                Expense.id != self.id
            ).order_by(Expense.date).all()
            
            all_expenses = [self] + related_expenses
            return sorted(all_expenses, key=lambda x: x.date)
        
        return []

class MonthlyGoal(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    year = db.Column(db.Integer, nullable=False)
    month = db.Column(db.Integer, nullable=False)
    target = db.Column(db.Integer, nullable=False, default=100)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    __table_args__ = (db.UniqueConstraint('year', 'month', name='_year_month_uc'),)
    
    def __repr__(self):
        return f'<MonthlyGoal {self.year}-{self.month}: {self.target}>'
    
    @classmethod
    def get_current_goal(cls):
        today = date.today()
        from flask import current_app
        goal = cls.query.filter_by(year=today.year, month=today.month).first()
        result = goal.target if goal else 100
        if current_app and current_app.logger and current_app.logger.isEnabledFor(10):
            current_app.logger.debug(
                f"MonthlyGoal.get_current_goal today={today} goal={'None' if goal is None else goal.target} result={result}")
        return result
    
    @classmethod
    def set_current_goal(cls, target):
        today = date.today()
        goal = cls.query.filter_by(year=today.year, month=today.month).first()
        
        if goal:
            goal.target = target
            goal.updated_at = datetime.utcnow()
        else:
            goal = cls(year=today.year, month=today.month, target=target)
            db.session.add(goal)
        
        db.session.commit()
        return goal

class DamagedProductLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=False)
    quantity = db.Column(db.Integer, nullable=False, default=1)
    purchase_price_snapshot = db.Column(db.Float, nullable=False, default=0.0)
    total_loss = db.Column(db.Float, nullable=False, default=0.0)
    return_order_item_id = db.Column(
        db.Integer,
        db.ForeignKey('return_order_item.id'),
        nullable=True,
        index=True
    )
    notes = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    created_by = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)

    created_by_employee = db.relationship('Employee', foreign_keys=[created_by], backref='damaged_product_logs')

    def __repr__(self):
        return f'<DamagedProductLog product_id={self.product_id} qty={self.quantity} loss={self.total_loss}>'


class Party(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, index=True)
    phone = db.Column(db.String(20), nullable=True)
    address = db.Column(db.String(200), nullable=True)
    notes = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    transactions = db.relationship('Transaction', backref='party', lazy=True, cascade="all, delete-orphan")
    
    @property
    def total_receivable(self):
        return sum(t.remaining_amount for t in self.transactions if t.transaction_type == 'receivable')
    
    @property
    def total_payable(self):
        return sum(t.remaining_amount for t in self.transactions if t.transaction_type == 'payable')
    
    @property
    def balance(self):
        return self.total_receivable - self.total_payable
    
    def __repr__(self):
        return f'<Party {self.name}>'


class Transaction(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    party_id = db.Column(db.Integer, db.ForeignKey('party.id'), nullable=False, index=True)
    transaction_type = db.Column(db.String(20), nullable=False, index=True)
    amount = db.Column(db.Float, nullable=False, default=0.0)
    paid_amount = db.Column(db.Float, nullable=False, default=0.0)
    category = db.Column(db.String(100), nullable=True)
    description = db.Column(db.Text, nullable=True)
    transaction_date = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    created_by = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)
    
    created_by_employee = db.relationship('Employee', foreign_keys=[created_by], backref='transactions_created')
    
    @property
    def remaining_amount(self):
        return (self.amount or 0) - (self.paid_amount or 0)
    
    def __repr__(self):
        return f'<Transaction {self.transaction_type} {self.amount} for party_id={self.party_id}>'


class TransactionPayment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    transaction_id = db.Column(db.Integer, db.ForeignKey('transaction.id', ondelete='CASCADE'), nullable=False, index=True)
    amount = db.Column(db.Float, nullable=False)
    notes = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    created_by = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)

    transaction = db.relationship('Transaction', backref=db.backref('payments', lazy='select'))
    created_by_employee = db.relationship('Employee', foreign_keys=[created_by])

    def __repr__(self):
        return f'<TransactionPayment tx={self.transaction_id} amount={self.amount}>'


class CapitalSnapshot(db.Model):
    __tablename__ = 'capital_snapshots'
    
    id = db.Column(db.Integer, primary_key=True)
    year = db.Column(db.Integer, nullable=False, index=True)
    month = db.Column(db.Integer, nullable=False, index=True)
    
    fixed_assets_value = db.Column(db.Float, nullable=False, default=0.0)
    pending_orders_value = db.Column(db.Float, nullable=False, default=0.0)
    stock_value = db.Column(db.Float, nullable=False, default=0.0)
    total_capital = db.Column(db.Float, nullable=False, default=0.0)
    total_debt = db.Column(db.Float, nullable=False, default=0.0)
    net_capital_after_debt = db.Column(db.Float, nullable=False, default=0.0)
    
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    __table_args__ = (db.UniqueConstraint('year', 'month', name='_capital_year_month_uc'),)
    
    def __repr__(self):
        return f'<CapitalSnapshot {self.year}-{self.month}: {self.total_capital} ج.م>'
    
    @classmethod
    def get_or_create_snapshot(cls, year, month):
        snapshot = cls.query.filter_by(year=year, month=month).first()
        if not snapshot:
            snapshot = cls(year=year, month=month)
            db.session.add(snapshot)
        return snapshot
    
    @classmethod
    def get_last_month_snapshot(cls):
        return cls.query.order_by(cls.year.desc(), cls.month.desc()).first()


class CapitalGrowthHistory(db.Model):
    """Capital snapshot history with growth rate."""
    __tablename__ = 'capital_growth_history'

    id = db.Column(db.Integer, primary_key=True)
    saved_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False, index=True)
    total_capital = db.Column(db.Float, nullable=False, default=0.0)
    previous_capital = db.Column(db.Float, nullable=True)
    growth_rate = db.Column(db.Float, nullable=True)   # Growth rate compared with the previous snapshot
    fixed_assets_value = db.Column(db.Float, nullable=False, default=0.0)
    stock_value = db.Column(db.Float, nullable=False, default=0.0)
    pending_orders_value = db.Column(db.Float, nullable=False, default=0.0)
    total_debt = db.Column(db.Float, nullable=False, default=0.0)
    net_capital_after_debt = db.Column(db.Float, nullable=False, default=0.0)

    def __repr__(self):
        return f'<CapitalGrowthHistory {self.saved_at}: {self.total_capital} ج.م ({self.growth_rate}%)>'

    @classmethod
    def get_history(cls, offset=0, limit=12):
        total = cls.query.count()
        items = cls.query.order_by(cls.saved_at.desc()).offset(offset).limit(limit).all()
        return items, total


class StockTake(db.Model):
    __tablename__ = 'stock_take'
    id = db.Column(db.Integer, primary_key=True)
    status = db.Column(db.String(20), nullable=False, default='in_progress')
    started_by_id = db.Column(db.Integer, db.ForeignKey('employee.id'))
    applied_by_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)
    discarded_by_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    applied_at = db.Column(db.DateTime, nullable=True)
    discarded_at = db.Column(db.DateTime, nullable=True)
    pdf_path = db.Column(db.String(255), nullable=True)
    items = db.relationship('StockTakeItem', backref='stocktake', cascade='all, delete-orphan')

    started_by = db.relationship('Employee', foreign_keys=[started_by_id])
    applied_by = db.relationship('Employee', foreign_keys=[applied_by_id])
    discarded_by = db.relationship('Employee', foreign_keys=[discarded_by_id])


class StockTakeItem(db.Model):
    __tablename__ = 'stock_take_item'
    id = db.Column(db.Integer, primary_key=True)
    stocktake_id = db.Column(db.Integer, db.ForeignKey('stock_take.id', ondelete='CASCADE'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id', ondelete='SET NULL'), nullable=True)
    product_name_snapshot = db.Column(db.String(100))
    system_stock = db.Column(db.Integer)
    counted_stock = db.Column(db.Integer, nullable=True)
    diff = db.Column(db.Integer)
    was_skipped = db.Column(db.Boolean, default=False)

    __table_args__ = (
        db.UniqueConstraint('stocktake_id', 'product_id', name='uq_stocktake_product'),
    )
