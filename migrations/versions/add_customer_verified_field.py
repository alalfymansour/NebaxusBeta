"""add_customer_verified_field

Revision ID: abc789verified123
Revises: xyz123capital456
Create Date: 2026-01-21 12:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'abc789verified123'
down_revision = 'xyz123capital456'
branch_labels = None
depends_on = None


def upgrade():
    # إضافة حقل customer_verified لجدول order
    op.add_column('order', sa.Column('customer_verified', sa.Boolean(), nullable=True, server_default='false'))
    # إضافة حقل customer_verified لجدول replacement_order
    op.add_column('replacement_order', sa.Column('customer_verified', sa.Boolean(), nullable=True, server_default='false'))


def downgrade():
    # إزالة الحقول
    op.drop_column('replacement_order', 'customer_verified')
    op.drop_column('order', 'customer_verified')
