"""add inventory_deducted column to order

Revision ID: add_inventory_deducted_to_order
Revises: 9a3d35e282b6
Create Date: 2025-09-27
"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'add_inventory_deducted_to_order'
down_revision = '9a3d35e282b6'
branch_labels = None
depends_on = None

def upgrade():
    bind = op.get_bind()
    inspector = sa.inspect(bind)
    existing_cols = [c['name'] for c in inspector.get_columns('order')]
    # إذا لم تُضف العمود بعد (فشل سابق جزئي) نضيفه وننشئ الفهرس
    if 'inventory_deducted' not in existing_cols:
        with op.batch_alter_table('order') as batch_op:
            batch_op.add_column(sa.Column('inventory_deducted', sa.Boolean(), nullable=False, server_default=sa.false()))
            # ملاحظة: create_index داخل batch يحتاج قائمة أعمدة
            batch_op.create_index('ix_order_inventory_deducted', ['inventory_deducted'])
        # إزالة القيمة الافتراضية بعد تعبئة القيم الحالية
        op.execute('ALTER TABLE "order" ALTER COLUMN inventory_deducted DROP DEFAULT')
    else:
        # العمود موجود بالفعل؛ نتأكد من وجود الفهرس
        existing_indexes = [ix['name'] for ix in inspector.get_indexes('order')]
        if 'ix_order_inventory_deducted' not in existing_indexes:
            op.create_index('ix_order_inventory_deducted', 'order', ['inventory_deducted'])


def downgrade():
    # حذف الفهرس أولاً لو موجود ثم العمود
    bind = op.get_bind()
    inspector = sa.inspect(bind)
    existing_indexes = [ix['name'] for ix in inspector.get_indexes('order')]
    if 'ix_order_inventory_deducted' in existing_indexes:
        op.drop_index('ix_order_inventory_deducted', table_name='order')
    existing_cols = [c['name'] for c in inspector.get_columns('order')]
    if 'inventory_deducted' in existing_cols:
        with op.batch_alter_table('order') as batch_op:
            batch_op.drop_column('inventory_deducted')
