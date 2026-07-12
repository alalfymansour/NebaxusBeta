"""add_is_returned_field_to_replacement_order_item

Revision ID: d7ca9a5e7daf
Revises: a9f1d2c3b4e5
Create Date: 2025-11-17 21:58:45.713562

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'd7ca9a5e7daf'
down_revision = 'a9f1d2c3b4e5'
branch_labels = None
depends_on = None


def upgrade():
    # إضافة حقل is_returned إلى جدول replacement_order_item
    op.add_column('replacement_order_item', sa.Column('is_returned', sa.Boolean(), nullable=True, server_default='false'))


def downgrade():
    # إزالة حقل is_returned
    op.drop_column('replacement_order_item', 'is_returned')
