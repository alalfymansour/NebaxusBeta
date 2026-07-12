"""drop can_view_reports and can_view_expected_profit columns

Revision ID: drop_reports_profit_cols
Revises: merge_heads_new_permissions_2026
Create Date: 2026-03-08

"""
from alembic import op
import sqlalchemy as sa

revision = 'drop_reports_profit_cols'
down_revision = 'merge_heads_new_permissions_2026'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('employee', schema=None) as batch_op:
        batch_op.drop_column('can_view_reports')
        batch_op.drop_column('can_view_expected_profit')


def downgrade():
    with op.batch_alter_table('employee', schema=None) as batch_op:
        batch_op.add_column(sa.Column('can_view_reports', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_expected_profit', sa.Boolean(), nullable=True, server_default=sa.false()))
