"""add capital growth history table

Revision ID: abc789growth123
Revises: xyz123capital456
Create Date: 2026-03-10 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'abc789growth123'
down_revision = 'xyz123capital456'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table('capital_growth_history',
    sa.Column('id', sa.Integer(), nullable=False),
    sa.Column('saved_at', sa.DateTime(), nullable=False),
    sa.Column('total_capital', sa.Float(), nullable=False),
    sa.Column('previous_capital', sa.Float(), nullable=True),
    sa.Column('growth_rate', sa.Float(), nullable=True),
    sa.Column('fixed_assets_value', sa.Float(), nullable=False),
    sa.Column('stock_value', sa.Float(), nullable=False),
    sa.Column('pending_orders_value', sa.Float(), nullable=False),
    sa.PrimaryKeyConstraint('id')
    )
    op.create_index(op.f('ix_capital_growth_history_saved_at'), 'capital_growth_history', ['saved_at'], unique=False)


def downgrade():
    op.drop_index(op.f('ix_capital_growth_history_saved_at'), table_name='capital_growth_history')
    op.drop_table('capital_growth_history')
