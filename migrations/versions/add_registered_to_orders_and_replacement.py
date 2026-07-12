"""add registered flag to orders and replacement_orders

Revision ID: add_registered_to_orders_and_replacement
Revises: b2c3d4e5f6a7
Create Date: 2025-12-29 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'add_registered_to_orders_and_replacement'
down_revision = 'b2c3d4e5f6a7'
branch_labels = None
depends_on = None


def upgrade():
    # Add registered column to order
    with op.batch_alter_table('order', schema=None) as batch_op:
        batch_op.add_column(sa.Column('registered', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.create_index('ix_order_registered', ['registered'])

    # Add registered column to replacement_order
    with op.batch_alter_table('replacement_order', schema=None) as batch_op:
        batch_op.add_column(sa.Column('registered', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.create_index('ix_replacement_order_registered', ['registered'])


def downgrade():
    with op.batch_alter_table('replacement_order', schema=None) as batch_op:
        batch_op.drop_index('ix_replacement_order_registered')
        batch_op.drop_column('registered')

    with op.batch_alter_table('order', schema=None) as batch_op:
        batch_op.drop_index('ix_order_registered')
        batch_op.drop_column('registered')
