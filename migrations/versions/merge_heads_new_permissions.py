"""merge heads: add_new_permissions_2026 and c2ba21cd8170

Revision ID: merge_heads_new_permissions_2026
Revises: add_new_permissions_2026, c2ba21cd8170
Create Date: 2026-03-08

"""
from alembic import op
import sqlalchemy as sa

revision = 'merge_heads_new_permissions_2026'
down_revision = ('add_new_permissions_2026', 'c2ba21cd8170')
branch_labels = None
depends_on = None


def upgrade():
    pass


def downgrade():
    pass
