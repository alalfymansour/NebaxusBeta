"""Add checkbox tracking timestamps and employee references

Revision ID: 9f8g7h6i5j4k_checkbox_tracking
Revises: add_customer_log_employee_id
Create Date: 2026-03-24 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '9f8g7h6i5j4k_checkbox_tracking'
down_revision = 'add_customer_log_employee_id'
branch_labels = None
depends_on = None


def upgrade():
    # Add columns to order table
    with op.batch_alter_table('order', schema=None) as batch_op:
        batch_op.add_column(sa.Column('customer_called_at', sa.DateTime(), nullable=True))
        batch_op.add_column(sa.Column('customer_called_by_id', sa.Integer(), nullable=True))
        batch_op.create_foreign_key('fk_order_customer_called_by_id', 'employee', ['customer_called_by_id'], ['id'], ondelete='SET NULL')
        
        batch_op.add_column(sa.Column('customer_verified_at', sa.DateTime(), nullable=True))
        batch_op.add_column(sa.Column('customer_verified_by_id', sa.Integer(), nullable=True))
        batch_op.create_foreign_key('fk_order_customer_verified_by_id', 'employee', ['customer_verified_by_id'], ['id'], ondelete='SET NULL')
        
        batch_op.add_column(sa.Column('registered_at', sa.DateTime(), nullable=True))
        batch_op.add_column(sa.Column('registered_by_id', sa.Integer(), nullable=True))
        batch_op.create_foreign_key('fk_order_registered_by_id', 'employee', ['registered_by_id'], ['id'], ondelete='SET NULL')

    # Add columns to replacement_order table
    with op.batch_alter_table('replacement_order', schema=None) as batch_op:
        batch_op.add_column(sa.Column('customer_called_at', sa.DateTime(), nullable=True))
        batch_op.add_column(sa.Column('customer_called_by_id', sa.Integer(), nullable=True))
        batch_op.create_foreign_key('fk_replacement_order_customer_called_by_id', 'employee', ['customer_called_by_id'], ['id'], ondelete='SET NULL')
        
        batch_op.add_column(sa.Column('customer_verified_at', sa.DateTime(), nullable=True))
        batch_op.add_column(sa.Column('customer_verified_by_id', sa.Integer(), nullable=True))
        batch_op.create_foreign_key('fk_replacement_order_customer_verified_by_id', 'employee', ['customer_verified_by_id'], ['id'], ondelete='SET NULL')
        
        batch_op.add_column(sa.Column('registered_at', sa.DateTime(), nullable=True))
        batch_op.add_column(sa.Column('registered_by_id', sa.Integer(), nullable=True))
        batch_op.create_foreign_key('fk_replacement_order_registered_by_id', 'employee', ['registered_by_id'], ['id'], ondelete='SET NULL')


def downgrade():
    # Remove columns from replacement_order table
    with op.batch_alter_table('replacement_order', schema=None) as batch_op:
        batch_op.drop_constraint('fk_replacement_order_registered_by_id', type_='foreignkey')
        batch_op.drop_column('registered_by_id')
        batch_op.drop_column('registered_at')
        
        batch_op.drop_constraint('fk_replacement_order_customer_verified_by_id', type_='foreignkey')
        batch_op.drop_column('customer_verified_by_id')
        batch_op.drop_column('customer_verified_at')
        
        batch_op.drop_constraint('fk_replacement_order_customer_called_by_id', type_='foreignkey')
        batch_op.drop_column('customer_called_by_id')
        batch_op.drop_column('customer_called_at')

    # Remove columns from order table
    with op.batch_alter_table('order', schema=None) as batch_op:
        batch_op.drop_constraint('fk_order_registered_by_id', type_='foreignkey')
        batch_op.drop_column('registered_by_id')
        batch_op.drop_column('registered_at')
        
        batch_op.drop_constraint('fk_order_customer_verified_by_id', type_='foreignkey')
        batch_op.drop_column('customer_verified_by_id')
        batch_op.drop_column('customer_verified_at')
        
        batch_op.drop_constraint('fk_order_customer_called_by_id', type_='foreignkey')
        batch_op.drop_column('customer_called_by_id')
        batch_op.drop_column('customer_called_at')
