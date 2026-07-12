"""add composite indexes

Revision ID: 7594286a35ce
Revises: 7b9e32ecfb73
Create Date: 2025-09-10 16:29:29.780663

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '7594286a35ce'
down_revision = '7b9e32ecfb73'
branch_labels = None
depends_on = None


def upgrade():
    # إنشاء الفهارس المركبة والجديدة
    # ملاحظة: لو كانت الجداول كبيرة ويمكنك تحمل زمن أطول بدون قفل كامل استخدم CONCURRENTLY يدوياً خارج Alembic
    op.create_index('ix_order_status_date', 'order', ['status', 'date'], unique=False)
    op.create_index('ix_replacement_order_status_date', 'replacement_order', ['status', 'date'], unique=False)
    op.create_index('ix_expense_category', 'expense', ['category'], unique=False)


def downgrade():
    op.drop_index('ix_expense_category', table_name='expense')
    op.drop_index('ix_replacement_order_status_date', table_name='replacement_order')
    op.drop_index('ix_order_status_date', table_name='order')
