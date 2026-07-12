"""add inventory_deducted column to replacement_order

Revision ID: a9f1d2c3b4e5
Revises: 8f55ebb84188
Create Date: 2025-10-05
"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'a9f1d2c3b4e5'
down_revision = '8f55ebb84188'
branch_labels = None
depends_on = None

def upgrade():
    bind = op.get_bind()
    inspector = sa.inspect(bind)
    existing_cols = [c['name'] for c in inspector.get_columns('replacement_order')]
    if 'inventory_deducted' not in existing_cols:
        with op.batch_alter_table('replacement_order') as batch_op:
            batch_op.add_column(sa.Column('inventory_deducted', sa.Boolean(), nullable=False, server_default=sa.false()))
            batch_op.create_index('ix_replacement_order_inventory_deducted', ['inventory_deducted'])
        op.execute('ALTER TABLE replacement_order ALTER COLUMN inventory_deducted DROP DEFAULT')
    else:
        existing_indexes = [ix['name'] for ix in inspector.get_indexes('replacement_order')]
        if 'ix_replacement_order_inventory_deducted' not in existing_indexes:
            op.create_index('ix_replacement_order_inventory_deducted', 'replacement_order', ['inventory_deducted'])


def downgrade():
    bind = op.get_bind()
    inspector = sa.inspect(bind)
    existing_indexes = [ix['name'] for ix in inspector.get_indexes('replacement_order')]
    if 'ix_replacement_order_inventory_deducted' in existing_indexes:
        op.drop_index('ix_replacement_order_inventory_deducted', table_name='replacement_order')
    existing_cols = [c['name'] for c in inspector.get_columns('replacement_order')]
    if 'inventory_deducted' in existing_cols:
        with op.batch_alter_table('replacement_order') as batch_op:
            batch_op.drop_column('inventory_deducted')
