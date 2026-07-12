"""add_customer_called_field

Revision ID: 654ad7d9890e
Revises: d7ca9a5e7daf
Create Date: 2025-11-17 22:35:00.677066

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '654ad7d9890e'
down_revision = 'd7ca9a5e7daf'
branch_labels = None
depends_on = None


def upgrade():
    # إضافة حقل customer_called لجدول order
    op.add_column('order', sa.Column('customer_called', sa.Boolean(), nullable=True, server_default='false'))
    # إضافة حقل customer_called لجدول replacement_order
    op.add_column('replacement_order', sa.Column('customer_called', sa.Boolean(), nullable=True, server_default='false'))


def downgrade():
    # إزالة الحقول
    op.drop_column('replacement_order', 'customer_called')
    op.drop_column('order', 'customer_called')
