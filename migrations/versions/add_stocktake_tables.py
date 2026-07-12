"""add stock_take and stock_take_item tables + can_manage_stocktake permission

Revision ID: add_stocktake_tables_001
Revises: add_order_type_walkin
Create Date: 2026-06-21 00:00:00.000000
"""
from alembic import op
import sqlalchemy as sa


revision = 'add_stocktake_tables_001'
down_revision = 'add_order_type_walkin'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('employee', sa.Column('can_manage_stocktake', sa.Boolean(), nullable=False, server_default='false'))

    op.create_table(
        'stock_take',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('status', sa.String(length=20), nullable=False, server_default='in_progress'),
        sa.Column('started_by_id', sa.Integer(), sa.ForeignKey('employee.id'), nullable=True),
        sa.Column('applied_by_id', sa.Integer(), sa.ForeignKey('employee.id'), nullable=True),
        sa.Column('discarded_by_id', sa.Integer(), sa.ForeignKey('employee.id'), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('applied_at', sa.DateTime(), nullable=True),
        sa.Column('discarded_at', sa.DateTime(), nullable=True),
        sa.Column('pdf_path', sa.String(length=255), nullable=True),
        sa.PrimaryKeyConstraint('id'),
    )

    op.create_table(
        'stock_take_item',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('stocktake_id', sa.Integer(), sa.ForeignKey('stock_take.id', ondelete='CASCADE'), nullable=False),
        sa.Column('product_id', sa.Integer(), sa.ForeignKey('product.id', ondelete='SET NULL'), nullable=True),
        sa.Column('product_name_snapshot', sa.String(length=100), nullable=True),
        sa.Column('system_stock', sa.Integer(), nullable=True),
        sa.Column('counted_stock', sa.Integer(), nullable=True),
        sa.Column('diff', sa.Integer(), nullable=True),
        sa.Column('was_skipped', sa.Boolean(), nullable=False, server_default='false'),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('stocktake_id', 'product_id', name='uq_stocktake_product'),
    )

    op.execute(
        "CREATE UNIQUE INDEX uq_one_in_progress_stocktake ON stock_take ((1)) WHERE status = 'in_progress'"
    )


def downgrade():
    op.execute("DROP INDEX IF EXISTS uq_one_in_progress_stocktake")
    op.drop_table('stock_take_item')
    op.drop_table('stock_take')
    op.drop_column('employee', 'can_manage_stocktake')
