"""Add purchase_price_snapshot to order_item

Revision ID: add_order_item_snapshot_20260507
Revises: rm_emp_pwd_cols_20260404
Create Date: 2026-05-07 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = 'add_order_item_snapshot_20260507'
down_revision = 'rm_emp_pwd_cols_20260404'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('order_item', schema=None) as batch_op:
        batch_op.add_column(sa.Column('purchase_price_snapshot', sa.Float(), nullable=True))

    op.execute(
        """
        UPDATE order_item
        SET purchase_price_snapshot = COALESCE(p.purchase_price, 0)
        FROM product p
        WHERE order_item.product_id = p.id
        """
    )
    op.execute(
        """
        UPDATE order_item
        SET purchase_price_snapshot = 0
        WHERE purchase_price_snapshot IS NULL
        """
    )


def downgrade():
    with op.batch_alter_table('order_item', schema=None) as batch_op:
        batch_op.drop_column('purchase_price_snapshot')
