"""add category to transaction

Revision ID: ghi345jkl678
Revises: def789ghi012
Create Date: 2025-12-12 00:02:00.000000

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'ghi345jkl678'
down_revision = 'def789ghi012'
branch_labels = None
depends_on = None


def upgrade():
    # إضافة عمود category للمعاملات
    op.add_column('transaction', sa.Column('category', sa.String(length=100), nullable=True))


def downgrade():
    # حذف العمود
    op.drop_column('transaction', 'category')
