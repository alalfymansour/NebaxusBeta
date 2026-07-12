"""merge heads: capital_growth_history and sold_products_by_qty

Revision ID: merge_growth_soldqty
Revises: abc789growth123, add_sold_products_by_qty
Create Date: 2026-03-10 13:00:00.000000

"""
from alembic import op
import sqlalchemy as sa

revision = 'merge_growth_soldqty'
down_revision = ('abc789growth123', 'add_sold_products_by_qty')
branch_labels = None
depends_on = None


def upgrade():
    pass


def downgrade():
    pass
