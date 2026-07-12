"""add submission_id to invoice

Revision ID: abcf853f094a
Revises: c2b28df36611
Create Date: 2026-06-23 16:53:00.000000

"""
from alembic import op
import sqlalchemy as sa
import uuid


revision = 'abcf853f094a'
down_revision = 'c2b28df36611'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('invoice', sa.Column('submission_id', sa.String(36), nullable=True))

    conn = op.get_bind()
    rows = conn.execute(sa.text("SELECT id FROM invoice WHERE submission_id IS NULL")).fetchall()
    for row in rows:
        conn.execute(
            sa.text("UPDATE invoice SET submission_id = :sid WHERE id = :id"),
            {"sid": str(uuid.uuid4()), "id": row[0]},
        )

    op.alter_column('invoice', 'submission_id', nullable=False)
    op.create_index('ix_invoice_submission_id', 'invoice', ['submission_id'], unique=True)


def downgrade():
    op.drop_index('ix_invoice_submission_id', table_name='invoice')
    op.drop_column('invoice', 'submission_id')
