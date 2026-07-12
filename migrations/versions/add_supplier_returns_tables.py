"""add supplier returns tables

Revision ID: a1b2c3d4e5f6
Revises: f1a2b3c4d5e6
Create Date: 2025-12-05 22:40:00.000000

"""
from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects import postgresql

# revision identifiers, used by Alembic.
revision = 'a1b2c3d4e5f6'
down_revision = 'f1a2b3c4d5e6'
branch_labels = None
depends_on = None


def upgrade():
    # إنشاء جدول supplier_returns
    op.create_table('supplier_returns',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('invoice_id', sa.Integer(), nullable=False),
        sa.Column('return_date', sa.DateTime(), nullable=True),
        sa.Column('total_amount', sa.Float(), nullable=True),
        sa.Column('notes', sa.Text(), nullable=True),
        sa.Column('created_by', sa.Integer(), nullable=True),
        sa.ForeignKeyConstraint(['created_by'], ['employee.id'], ),
        sa.ForeignKeyConstraint(['invoice_id'], ['invoice.id'], ),
        sa.PrimaryKeyConstraint('id')
    )
    op.create_index(op.f('ix_supplier_returns_invoice_id'), 'supplier_returns', ['invoice_id'], unique=False)
    op.create_index(op.f('ix_supplier_returns_return_date'), 'supplier_returns', ['return_date'], unique=False)

    # إنشاء جدول supplier_return_items
    op.create_table('supplier_return_items',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('return_id', sa.Integer(), nullable=False),
        sa.Column('invoice_item_id', sa.Integer(), nullable=False),
        sa.Column('product_id', sa.Integer(), nullable=False),
        sa.Column('quantity', sa.Integer(), nullable=False),
        sa.Column('price', sa.Float(), nullable=False),
        sa.Column('total', sa.Float(), nullable=False),
        sa.ForeignKeyConstraint(['invoice_item_id'], ['invoice_item.id'], ),
        sa.ForeignKeyConstraint(['product_id'], ['product.id'], ),
        sa.ForeignKeyConstraint(['return_id'], ['supplier_returns.id'], ),
        sa.PrimaryKeyConstraint('id')
    )
    op.create_index(op.f('ix_supplier_return_items_return_id'), 'supplier_return_items', ['return_id'], unique=False)


def downgrade():
    # حذف الجداول بالترتيب العكسي
    op.drop_index(op.f('ix_supplier_return_items_return_id'), table_name='supplier_return_items')
    op.drop_table('supplier_return_items')
    
    op.drop_index(op.f('ix_supplier_returns_return_date'), table_name='supplier_returns')
    op.drop_index(op.f('ix_supplier_returns_invoice_id'), table_name='supplier_returns')
    op.drop_table('supplier_returns')
