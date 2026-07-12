"""add draft fields to replacement order

Revision ID: f9a8b7c6d5e4
Revises: 45da718cbf20
Create Date: 2026-02-22 16:20:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'f9a8b7c6d5e4'
down_revision = '45da718cbf20'
branch_labels = None
depends_on = None


def upgrade():
    # Add draft fields to replacement_order table
    op.add_column('replacement_order', sa.Column('is_draft', sa.Boolean(), nullable=False, server_default='0'))
    op.add_column('replacement_order', sa.Column('draft_step', sa.String(length=50), nullable=True))
    op.add_column('replacement_order', sa.Column('draft_data', sa.Text(), nullable=True))
    
    # Create indexes
    op.create_index(op.f('ix_replacement_order_is_draft'), 'replacement_order', ['is_draft'], unique=False)


def downgrade():
    # Drop indexes
    op.drop_index(op.f('ix_replacement_order_is_draft'), table_name='replacement_order')
    
    # Remove columns
    op.drop_column('replacement_order', 'draft_data')
    op.drop_column('replacement_order', 'draft_step')
    op.drop_column('replacement_order', 'is_draft')
