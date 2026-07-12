"""add customer snapshot columns to order table

Revision ID: add_customer_snapshot_order
Revises: add_new_permissions_2026
Create Date: 2026-03-11

"""
from alembic import op
import sqlalchemy as sa

revision = 'add_customer_snapshot_order'
down_revision = 'd9e0f1a2b3c4'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('order', schema=None) as batch_op:
        batch_op.add_column(sa.Column('customer_name', sa.String(100), nullable=True))
        batch_op.add_column(sa.Column('customer_phone', sa.String(20), nullable=True))
        batch_op.add_column(sa.Column('customer_governorate', sa.String(100), nullable=True))
        batch_op.add_column(sa.Column('customer_address_details', sa.String(200), nullable=True))


def downgrade():
    with op.batch_alter_table('order', schema=None) as batch_op:
        batch_op.drop_column('customer_address_details')
        batch_op.drop_column('customer_governorate')
        batch_op.drop_column('customer_phone')
        batch_op.drop_column('customer_name')
