"""add party and transaction tables

Revision ID: abc123def456
Revises: a1b2c3d4e5f6
Create Date: 2025-12-12 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'abc123def456'
down_revision = 'a1b2c3d4e5f6'
branch_labels = None
depends_on = None


def upgrade():
    # إنشاء جدول party (المتعاملين)
    op.create_table('party',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('name', sa.String(length=100), nullable=False),
        sa.Column('phone', sa.String(length=20), nullable=True),
        sa.Column('address', sa.String(length=200), nullable=True),
        sa.Column('notes', sa.Text(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('updated_at', sa.DateTime(), nullable=True),
        sa.PrimaryKeyConstraint('id')
    )
    op.create_index(op.f('ix_party_name'), 'party', ['name'], unique=False)

    # إنشاء جدول transaction (المعاملات)
    op.create_table('transaction',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('party_id', sa.Integer(), nullable=False),
        sa.Column('transaction_type', sa.String(length=20), nullable=False),
        sa.Column('amount', sa.Float(), nullable=False),
        sa.Column('description', sa.Text(), nullable=True),
        sa.Column('transaction_date', sa.DateTime(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('created_by', sa.Integer(), nullable=True),
        sa.ForeignKeyConstraint(['created_by'], ['employee.id'], ),
        sa.ForeignKeyConstraint(['party_id'], ['party.id'], ),
        sa.PrimaryKeyConstraint('id')
    )
    op.create_index(op.f('ix_transaction_party_id'), 'transaction', ['party_id'], unique=False)
    op.create_index(op.f('ix_transaction_transaction_date'), 'transaction', ['transaction_date'], unique=False)
    op.create_index(op.f('ix_transaction_transaction_type'), 'transaction', ['transaction_type'], unique=False)


def downgrade():
    # حذف الجداول
    op.drop_index(op.f('ix_transaction_transaction_type'), table_name='transaction')
    op.drop_index(op.f('ix_transaction_transaction_date'), table_name='transaction')
    op.drop_index(op.f('ix_transaction_party_id'), table_name='transaction')
    op.drop_table('transaction')
    
    op.drop_index(op.f('ix_party_name'), table_name='party')
    op.drop_table('party')
