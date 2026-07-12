"""remove can_clear_database column from employee

Revision ID: remove_can_clear_database
Revises: 42dbddd2252d
Create Date: 2025-09-06
"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'remove_can_clear_database'
down_revision = 'ef3c6090b453'
branch_labels = None
depends_on = None

def upgrade():
    with op.batch_alter_table('employee') as batch_op:
        try:
            batch_op.drop_column('can_clear_database')
        except Exception:
            pass


def downgrade():
    with op.batch_alter_table('employee') as batch_op:
        batch_op.add_column(sa.Column('can_clear_database', sa.Boolean(), nullable=True))
