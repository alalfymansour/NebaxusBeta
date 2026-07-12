"""merge heads after inventory_deducted addition

Revision ID: 8f55ebb84188
Revises: 1b2c3d4e5f60, add_inventory_deducted_to_order
Create Date: 2025-09-27 22:45:35.670992

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '8f55ebb84188'
down_revision = ('1b2c3d4e5f60', 'add_inventory_deducted_to_order')
branch_labels = None
depends_on = None


def upgrade():
    pass


def downgrade():
    pass
