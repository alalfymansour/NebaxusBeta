"""add bundle product support

Revision ID: add_bundle_product_support
Revises: 654ad7d9890e
Create Date: 2025-11-29 13:20:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'add_bundle_product_support'
down_revision = '654ad7d9890e'
branch_labels = None
depends_on = None


def upgrade():
    # Add is_bundle column to product table
    with op.batch_alter_table('product', schema=None) as batch_op:
        batch_op.add_column(sa.Column('is_bundle', sa.Boolean(), nullable=False, server_default='false'))
    
    # Create bundle_item table
    op.create_table('bundle_item',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('bundle_id', sa.Integer(), nullable=False),
        sa.Column('product_id', sa.Integer(), nullable=False),
        sa.Column('sale_price_in_bundle', sa.Float(), nullable=False, server_default='0'),
        sa.ForeignKeyConstraint(['bundle_id'], ['product.id'], ondelete='CASCADE'),
        sa.ForeignKeyConstraint(['product_id'], ['product.id'], ondelete='CASCADE'),
        sa.PrimaryKeyConstraint('id')
    )


def downgrade():
    # Drop bundle_item table
    op.drop_table('bundle_item')
    
    # Remove is_bundle column from product table
    with op.batch_alter_table('product', schema=None) as batch_op:
        batch_op.drop_column('is_bundle')
