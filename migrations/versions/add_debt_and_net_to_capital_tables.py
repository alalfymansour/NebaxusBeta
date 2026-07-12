"""Add debt/net fields to capital snapshot and growth history

Revision ID: add_capital_debt_net_2026
Revises: 9f8g7h6i5j4k_checkbox_tracking
Create Date: 2026-04-04 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'add_capital_debt_net_2026'
down_revision = '9f8g7h6i5j4k_checkbox_tracking'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('capital_snapshots', schema=None) as batch_op:
        batch_op.add_column(sa.Column('total_debt', sa.Float(), nullable=False, server_default=sa.text('0')))
        batch_op.add_column(sa.Column('net_capital_after_debt', sa.Float(), nullable=False, server_default=sa.text('0')))

    with op.batch_alter_table('capital_growth_history', schema=None) as batch_op:
        batch_op.add_column(sa.Column('total_debt', sa.Float(), nullable=False, server_default=sa.text('0')))
        batch_op.add_column(sa.Column('net_capital_after_debt', sa.Float(), nullable=False, server_default=sa.text('0')))


def downgrade():
    with op.batch_alter_table('capital_growth_history', schema=None) as batch_op:
        batch_op.drop_column('net_capital_after_debt')
        batch_op.drop_column('total_debt')

    with op.batch_alter_table('capital_snapshots', schema=None) as batch_op:
        batch_op.drop_column('net_capital_after_debt')
        batch_op.drop_column('total_debt')
