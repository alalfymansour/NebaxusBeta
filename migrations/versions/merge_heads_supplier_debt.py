"""merge heads for supplier debt feature

Revision ID: merge_supplier_debt
Revises: remove_can_clear_database, a1b2c3d4e5f6
Create Date: 2025-12-06 12:32:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'merge_supplier_debt'
down_revision = ('remove_can_clear_database', 'a1b2c3d4e5f6')
branch_labels = None
depends_on = None


def upgrade():
    pass


def downgrade():
    pass
