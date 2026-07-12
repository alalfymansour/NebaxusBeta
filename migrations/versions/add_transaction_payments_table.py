"""add transaction payments table

Revision ID: add_transaction_payments_table
Revises: 
Create Date: 2025-12-17 00:00:00.000000
"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'add_transaction_payments_table'
down_revision = 'mno789pqr012'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'transaction_payment',
        sa.Column('id', sa.Integer(), primary_key=True),
        sa.Column('transaction_id', sa.Integer(), sa.ForeignKey('transaction.id', ondelete='CASCADE'), nullable=False, index=True),
        sa.Column('amount', sa.Float(), nullable=False),
        sa.Column('notes', sa.Text(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.Column('created_by', sa.Integer(), sa.ForeignKey('employee.id'), nullable=True),
    )


def downgrade():
    op.drop_table('transaction_payment')
