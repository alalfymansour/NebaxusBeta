"""add wassalha fields to order and replacement_order

Revision ID: add_wassalha_fields
Revises: add_attendance_system
Create Date: 2026-07-10
"""
from alembic import op
import sqlalchemy as sa


revision = 'add_wassalha_fields'
down_revision = 'add_attendance_system'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('order', sa.Column('weight',         sa.Float(),    nullable=True))
    op.add_column('order', sa.Column('package_volume', sa.String(20), nullable=True))
    op.add_column('order', sa.Column('delivery_notes', sa.Text(),     nullable=True))

    op.add_column('replacement_order', sa.Column('weight',         sa.Float(),    nullable=True))
    op.add_column('replacement_order', sa.Column('package_volume', sa.String(20), nullable=True))
    op.add_column('replacement_order', sa.Column('delivery_notes', sa.Text(),     nullable=True))


def downgrade():
    op.drop_column('order', 'delivery_notes')
    op.drop_column('order', 'package_volume')
    op.drop_column('order', 'weight')

    op.drop_column('replacement_order', 'delivery_notes')
    op.drop_column('replacement_order', 'package_volume')
    op.drop_column('replacement_order', 'weight')
