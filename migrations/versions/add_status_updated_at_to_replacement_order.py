"""add_status_updated_at_to_replacement_order

Revision ID: f5g6h7i8j9k0
Revises: e8f9a1b2c3d4
Create Date: 2025-12-06 14:55:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'f5g6h7i8j9k0'
down_revision = 'e8f9a1b2c3d4'
branch_labels = None
depends_on = None


def upgrade():
    # إضافة حقل status_updated_at لجدول replacement_order
    op.add_column('replacement_order', sa.Column('status_updated_at', sa.DateTime(), nullable=True))
    
    # تحديث القيم الحالية: نسخ قيمة date إلى status_updated_at لطلبات الاستبدال الموجودة
    op.execute("UPDATE replacement_order SET status_updated_at = date WHERE status_updated_at IS NULL")
    
    # إضافة index للحقل الجديد
    op.create_index(op.f('ix_replacement_order_status_updated_at'), 'replacement_order', ['status_updated_at'], unique=False)


def downgrade():
    # إزالة الحقل والـ index
    op.drop_index(op.f('ix_replacement_order_status_updated_at'), table_name='replacement_order')
    op.drop_column('replacement_order', 'status_updated_at')
