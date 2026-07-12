"""add original_replacement_order_id to replacement_order

Revision ID: c7d8e9f0a1b2
Revises: f9a8b7c6d5e4
Create Date: 2026-03-10 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'c7d8e9f0a1b2'
down_revision = 'f9a8b7c6d5e4'
branch_labels = None
depends_on = None


def upgrade():
    conn = op.get_bind()

    # Add column only if it doesn't already exist
    result = conn.execute(sa.text("""
        SELECT column_name FROM information_schema.columns
        WHERE table_name='replacement_order'
        AND column_name='original_replacement_order_id'
    """))
    if not result.fetchone():
        op.add_column('replacement_order',
            sa.Column('original_replacement_order_id', sa.Integer(), nullable=True)
        )

    # Add FK only if it doesn't already exist
    result = conn.execute(sa.text("""
        SELECT constraint_name FROM information_schema.table_constraints
        WHERE table_name='replacement_order'
        AND constraint_name='fk_replacement_order_original_replacement'
    """))
    if not result.fetchone():
        op.create_foreign_key(
            'fk_replacement_order_original_replacement',
            'replacement_order', 'replacement_order',
            ['original_replacement_order_id'], ['id'],
            ondelete='SET NULL'
        )

    # Add index only if it doesn't already exist
    conn.execute(sa.text("""
        CREATE INDEX IF NOT EXISTS ix_replacement_order_original_replacement_order_id
        ON replacement_order (original_replacement_order_id)
    """))


def downgrade():
    op.drop_index('ix_replacement_order_original_replacement_order_id', table_name='replacement_order')
    op.drop_constraint('fk_replacement_order_original_replacement', 'replacement_order', type_='foreignkey')
    op.drop_column('replacement_order', 'original_replacement_order_id')