"""add paid_amount to transaction

Revision ID: jkl456mno789
Revises: ghi345jkl678
Create Date: 2025-12-12 03:55:00.000000

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'jkl456mno789'
down_revision = 'ghi345jkl678'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('transaction', sa.Column('paid_amount', sa.Float(), nullable=False, server_default='0.0'))
    with op.batch_alter_table('transaction', schema=None) as batch_op:
        batch_op.alter_column('paid_amount', server_default=None)


def downgrade():
    op.drop_column('transaction', 'paid_amount')
