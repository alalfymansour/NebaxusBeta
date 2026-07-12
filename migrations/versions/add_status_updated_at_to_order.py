"""add_status_updated_at_to_order

Revision ID: e8f9a1b2c3d4
Revises: 654ad7d9890e
Create Date: 2025-12-05 10:30:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'e8f9a1b2c3d4'
down_revision = '654ad7d9890e'
branch_labels = None
depends_on = None


def upgrade():
    # إضافة حقل status_updated_at لجدول order
    op.add_column('order', sa.Column('status_updated_at', sa.DateTime(), nullable=True))
    
    # تحديث القيم الحالية: نسخ قيمة date إلى status_updated_at للطلبات الموجودة
    op.execute("UPDATE \"order\" SET status_updated_at = date WHERE status_updated_at IS NULL")
    
    # إضافة index للحقل الجديد
    op.create_index(op.f('ix_order_status_updated_at'), 'order', ['status_updated_at'], unique=False)


def downgrade():
    # إزالة الحقل والـ index
    op.drop_index(op.f('ix_order_status_updated_at'), table_name='order')
    op.drop_column('order', 'status_updated_at')
