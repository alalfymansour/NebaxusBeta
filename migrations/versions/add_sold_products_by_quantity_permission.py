"""add can_view_sold_products_by_quantity permission to employee

Revision ID: add_sold_products_by_qty
Revises: drop_reports_profit_cols
Create Date: 2026-03-08

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'add_sold_products_by_qty'
down_revision = 'drop_reports_profit_cols'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('employee', sa.Column('can_view_sold_products_by_quantity', sa.Boolean(), nullable=True))
    op.execute("UPDATE employee SET can_view_sold_products_by_quantity = FALSE WHERE can_view_sold_products_by_quantity IS NULL")
    op.alter_column('employee', 'can_view_sold_products_by_quantity', nullable=False)


def downgrade():
    op.drop_column('employee', 'can_view_sold_products_by_quantity')
