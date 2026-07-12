"""add order_edit_log table for tracking order edits with employee info

Revision ID: add_order_edit_log
Revises: add_customer_snapshot_order
Create Date: 2026-03-11

"""
from alembic import op
import sqlalchemy as sa

revision = 'add_order_edit_log'
down_revision = 'add_customer_snapshot_order'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'order_edit_log',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('order_id', sa.Integer(), nullable=True),
        sa.Column('replacement_order_id', sa.Integer(), nullable=True),
        sa.Column('employee_id', sa.Integer(), nullable=True),
        sa.Column('timestamp', sa.DateTime(), nullable=False),
        sa.Column('description', sa.Text(), nullable=False),
        sa.ForeignKeyConstraint(['employee_id'], ['employee.id'], ondelete='SET NULL'),
        sa.ForeignKeyConstraint(['order_id'], ['order.id'], ondelete='CASCADE'),
        sa.ForeignKeyConstraint(['replacement_order_id'], ['replacement_order.id'], ondelete='CASCADE'),
        sa.PrimaryKeyConstraint('id')
    )
    op.create_index('ix_order_edit_log_order_id', 'order_edit_log', ['order_id'])
    op.create_index('ix_order_edit_log_replacement_order_id', 'order_edit_log', ['replacement_order_id'])
    op.create_index('ix_order_edit_log_timestamp', 'order_edit_log', ['timestamp'])


def downgrade():
    op.drop_index('ix_order_edit_log_timestamp', table_name='order_edit_log')
    op.drop_index('ix_order_edit_log_replacement_order_id', table_name='order_edit_log')
    op.drop_index('ix_order_edit_log_order_id', table_name='order_edit_log')
    op.drop_table('order_edit_log')
