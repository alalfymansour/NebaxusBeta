"""add is_nearest_post_branch to order

Revision ID: add_nearest_post_branch
Revises: mno789pqr012
Create Date: 2026-02-01

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'add_nearest_post_branch'
down_revision = 'mno789pqr012'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('order', schema=None) as batch_op:
        batch_op.add_column(sa.Column('is_nearest_post_branch', sa.Boolean(), nullable=True, server_default='0'))


def downgrade():
    with op.batch_alter_table('order', schema=None) as batch_op:
        batch_op.drop_column('is_nearest_post_branch')
