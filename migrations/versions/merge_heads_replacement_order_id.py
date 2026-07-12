"""merge heads: add_original_replacement_order_id and merge_growth_soldqty

Revision ID: d9e0f1a2b3c4
Revises: c7d8e9f0a1b2, merge_growth_soldqty
Create Date: 2026-03-10 19:00:00.000000

"""
from alembic import op
import sqlalchemy as sa

revision = 'd9e0f1a2b3c4'
down_revision = ('c7d8e9f0a1b2', 'merge_growth_soldqty')
branch_labels = None
depends_on = None


def upgrade():
    pass


def downgrade():
    pass
