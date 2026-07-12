"""add alternative customer data to replacement order

Revision ID: add_alt_customer_data_repl
Revises: 13975594f5a5
Create Date: 2026-02-07 12:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'add_alt_customer_data_repl'
down_revision = '13975594f5a5'
branch_labels = None
depends_on = None


def upgrade():
    # إضافة حقول البيانات البديلة لطلبات الاستبدال
    op.add_column('replacement_order', sa.Column('alternative_name', sa.String(length=200), nullable=True))
    op.add_column('replacement_order', sa.Column('alternative_phone', sa.String(length=20), nullable=True))
    op.add_column('replacement_order', sa.Column('alternative_governorate', sa.String(length=100), nullable=True))
    op.add_column('replacement_order', sa.Column('alternative_address_details', sa.Text(), nullable=True))


def downgrade():
    # إزالة حقول البيانات البديلة
    op.drop_column('replacement_order', 'alternative_address_details')
    op.drop_column('replacement_order', 'alternative_governorate')
    op.drop_column('replacement_order', 'alternative_phone')
    op.drop_column('replacement_order', 'alternative_name')
