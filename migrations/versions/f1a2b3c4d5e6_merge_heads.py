"""merge_heads

Revision ID: f1a2b3c4d5e6
Revises: 612c49e763e9, e8f9a1b2c3d4
Create Date: 2025-12-05 08:40:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'f1a2b3c4d5e6'
down_revision = ('612c49e763e9', 'e8f9a1b2c3d4')
branch_labels = None
depends_on = None


def upgrade():
    # No changes needed - just merging heads
    pass


def downgrade():
    # No changes needed - just merging heads
    pass
