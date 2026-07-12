"""add_order_type_to_orders_and_make_customer_id_nullable

Revision ID: add_order_type_walkin
Revises: add_order_item_snapshot_20260507
Create Date: 2026-06-20 13:30:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = 'add_order_type_walkin'
down_revision = 'add_order_item_snapshot_20260507'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column(
        'order',
        sa.Column('order_type', sa.String(20), nullable=False, server_default='delivery'),
    )
    op.alter_column('order', 'customer_id', existing_type=sa.Integer(), nullable=True)

    # allowed order_type values: 'delivery', 'walkin'.
    # walkin orders must never have a customer_id; delivery orders may or may not.
    op.create_check_constraint(
        'order_type_customer_check',
        'order',
        "(order_type = 'walkin' AND customer_id IS NULL) OR (order_type = 'delivery')",
    )


def downgrade():
    op.drop_constraint('order_type_customer_check', 'order', type_='check')
    op.alter_column('order', 'customer_id', existing_type=sa.Integer(), nullable=False)
    op.drop_column('order', 'order_type')
