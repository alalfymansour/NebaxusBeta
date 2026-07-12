"""merge heads: status_updated_at and supplier_debt

Revision ID: m1n2o3p4q5r6
Revises: f5g6h7i8j9k0, add_supplier_debt_001
Create Date: 2025-12-06 15:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'm1n2o3p4q5r6'
down_revision = ('f5g6h7i8j9k0', 'add_supplier_debt_001')
branch_labels = None
depends_on = None


def upgrade():
    # This is a merge migration - no changes needed
    pass


def downgrade():
    # This is a merge migration - no changes needed
    pass
