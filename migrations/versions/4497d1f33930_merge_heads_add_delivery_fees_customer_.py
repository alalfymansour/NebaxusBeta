"""merge heads: add_delivery_fees_customer + a535e6d5da70

Revision ID: 4497d1f33930
Revises: a535e6d5da70, add_delivery_fees_customer
Create Date: 2025-08-21 17:16:37.084663

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '4497d1f33930'
down_revision = ('a535e6d5da70', 'add_delivery_fees_customer')
branch_labels = None
depends_on = None


def upgrade():
    pass


def downgrade():
    pass
