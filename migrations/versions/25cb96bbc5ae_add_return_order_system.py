"""add_return_order_system

Revision ID: 25cb96bbc5ae
Revises: abcf853f094a
Create Date: 2026-06-24 15:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = '25cb96bbc5ae'
down_revision = 'abcf853f094a'
branch_labels = None
depends_on = None


def upgrade():
    # Create return_order table
    op.create_table('return_order',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('original_order_id', sa.Integer(), nullable=True),
        sa.Column('original_replacement_order_id', sa.Integer(), nullable=True),
        sa.Column('customer_id', sa.Integer(), nullable=False),
        sa.Column('customer_refund_amount', sa.Float(), nullable=False, server_default='0.0'),
        sa.Column('received_by_id', sa.Integer(), nullable=False),
        sa.Column('created_at', sa.DateTime(), nullable=False, server_default=sa.text('CURRENT_TIMESTAMP')),
        sa.ForeignKeyConstraint(['original_order_id'], ['order.id'], ),
        sa.ForeignKeyConstraint(['original_replacement_order_id'], ['replacement_order.id'], ),
        sa.ForeignKeyConstraint(['customer_id'], ['customer.id'], ),
        sa.ForeignKeyConstraint(['received_by_id'], ['employee.id'], ),
        sa.CheckConstraint('(original_order_id IS NOT NULL)::int + (original_replacement_order_id IS NOT NULL)::int = 1', name='ck_return_order_single_source'),
        sa.PrimaryKeyConstraint('id')
    )
    op.create_index('ix_return_order_customer_created', 'return_order', ['customer_id', 'created_at'])
    op.create_index('ix_return_order_source_order', 'return_order', ['original_order_id'])
    op.create_index('ix_return_order_source_replacement', 'return_order', ['original_replacement_order_id'])

    # Create return_order_item table
    op.create_table('return_order_item',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('return_order_id', sa.Integer(), nullable=False),
        sa.Column('product_id', sa.Integer(), nullable=False),
        sa.Column('quantity', sa.Integer(), nullable=False, server_default='1'),
        sa.Column('item_condition', sa.String(20), nullable=False),
        sa.Column('return_reason', sa.String(50), nullable=False),
        sa.Column('unit_sale_price_snapshot', sa.Float(), nullable=False),
        sa.Column('unit_purchase_price_snapshot', sa.Float(), nullable=False),
        sa.Column('inspected_by_id', sa.Integer(), nullable=False),
        sa.Column('inspected_at', sa.DateTime(), nullable=False, server_default=sa.text('CURRENT_TIMESTAMP')),
        sa.ForeignKeyConstraint(['return_order_id'], ['return_order.id'], ),
        sa.ForeignKeyConstraint(['product_id'], ['product.id'], ),
        sa.ForeignKeyConstraint(['inspected_by_id'], ['employee.id'], ),
        sa.PrimaryKeyConstraint('id')
    )
    op.create_index('ix_return_order_item_product', 'return_order_item', ['product_id', 'return_order_id'])

    # Add return_order_id to customer_log
    op.add_column('customer_log', sa.Column('return_order_id', sa.Integer(), nullable=True))
    op.create_index(op.f('ix_customer_log_return_order_id'), 'customer_log', ['return_order_id'])
    op.create_foreign_key('fk_customer_log_return_order_id', 'customer_log', 'return_order', ['return_order_id'], ['id'])

    # Add return_order_item_id to damaged_product_log
    op.add_column('damaged_product_log', sa.Column('return_order_item_id', sa.Integer(), nullable=True))
    op.create_index(op.f('ix_damaged_product_log_return_order_item_id'), 'damaged_product_log', ['return_order_item_id'])
    op.create_foreign_key('fk_damaged_product_log_return_order_item_id', 'damaged_product_log', 'return_order_item', ['return_order_item_id'], ['id'])

    # Add return permission columns to employee
    op.add_column('employee', sa.Column('can_view_returns', sa.Boolean(), nullable=False, server_default=sa.text('false')))
    op.add_column('employee', sa.Column('can_add_returns', sa.Boolean(), nullable=False, server_default=sa.text('false')))
    op.add_column('employee', sa.Column('can_view_returns_by_state', sa.Boolean(), nullable=False, server_default=sa.text('false')))
    op.add_column('employee', sa.Column('can_delete_returns', sa.Boolean(), nullable=False, server_default=sa.text('false')))


def downgrade():
    op.drop_column('employee', 'can_delete_returns')
    op.drop_column('employee', 'can_view_returns_by_state')
    op.drop_column('employee', 'can_add_returns')
    op.drop_column('employee', 'can_view_returns')

    op.drop_constraint('fk_damaged_product_log_return_order_item_id', 'damaged_product_log', type_='foreignkey')
    op.drop_index(op.f('ix_damaged_product_log_return_order_item_id'), table_name='damaged_product_log')
    op.drop_column('damaged_product_log', 'return_order_item_id')

    op.drop_constraint('fk_customer_log_return_order_id', 'customer_log', type_='foreignkey')
    op.drop_index(op.f('ix_customer_log_return_order_id'), table_name='customer_log')
    op.drop_column('customer_log', 'return_order_id')

    op.drop_index('ix_return_order_item_product', table_name='return_order_item')
    op.drop_table('return_order_item')

    op.drop_index('ix_return_order_source_replacement', table_name='return_order')
    op.drop_index('ix_return_order_source_order', table_name='return_order')
    op.drop_index('ix_return_order_customer_created', table_name='return_order')
    op.drop_table('return_order')
