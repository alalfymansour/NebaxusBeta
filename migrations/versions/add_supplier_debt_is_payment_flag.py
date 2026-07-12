"""add is_payment flag to supplier_debts

Revision ID: b2c3d4e5f6a7
Revises: add_transaction_payments_table
Create Date: 2025-12-17 17:15:00.000000
"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'b2c3d4e5f6a7'
down_revision = 'add_transaction_payments_table'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('supplier_debts') as batch_op:
        batch_op.add_column(sa.Column('is_payment', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.create_index('ix_supplier_debts_is_payment', ['is_payment'])

    # Optional data migration: mark existing records that look like payments
    op.execute(
        """
        UPDATE supplier_debts
        SET is_payment = true
        WHERE COALESCE(amount,0) = 0 AND COALESCE(paid_amount,0) > 0
        """
    )


def downgrade():
    with op.batch_alter_table('supplier_debts') as batch_op:
        batch_op.drop_index('ix_supplier_debts_is_payment')
        batch_op.drop_column('is_payment')
