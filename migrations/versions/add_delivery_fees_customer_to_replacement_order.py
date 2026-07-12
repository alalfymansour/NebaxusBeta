"""add delivery_fees_customer to replacement_order

Revision ID: add_delivery_fees_customer
Revises: 9a3d35e282b6
Create Date: 2025-08-21

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'add_delivery_fees_customer'
down_revision = '9a3d35e282b6'
branch_labels = None
depends_on = None

def upgrade():
    with op.batch_alter_table('replacement_order', schema=None) as batch_op:
        batch_op.add_column(sa.Column('delivery_fees_customer', sa.Float(), nullable=True))


def downgrade():
    with op.batch_alter_table('replacement_order', schema=None) as batch_op:
        batch_op.drop_column('delivery_fees_customer')
