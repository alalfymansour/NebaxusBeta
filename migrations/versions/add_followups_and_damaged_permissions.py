"""add followups and damaged products permission fields

Revision ID: 1b2c3d4e5f60
Revises: 7594286a35ce  # (updated from original 698205474ba0 before applying)
Create Date: 2025-09-15
"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '1b2c3d4e5f60'
# NOTE: Updated to attach this migration to the latest known head to avoid creating a new branch / multiple heads.
# Previous value was '698205474ba0' which is an older revision in the graph; changing because migration not yet applied.
down_revision = '7594286a35ce'
branch_labels = None
depends_on = None

def upgrade():
    # نضيف الأعمدة بقيمة افتراضية FALSE على مستوى السيرفر لتعبئة الصفوف الحالية دون الحاجة لأمر UPDATE منفصل
    # (PostgreSQL لا يقبل تعيين 0 مباشرة إلى Boolean بدون cast)  
    with op.batch_alter_table('employee') as batch_op:
        batch_op.add_column(sa.Column('can_view_followups', sa.Boolean(), server_default=sa.text('false'), nullable=False))
        batch_op.add_column(sa.Column('can_add_followups', sa.Boolean(), server_default=sa.text('false'), nullable=False))
        batch_op.add_column(sa.Column('can_edit_followups', sa.Boolean(), server_default=sa.text('false'), nullable=False))
        batch_op.add_column(sa.Column('can_delete_followups', sa.Boolean(), server_default=sa.text('false'), nullable=False))
        batch_op.add_column(sa.Column('can_view_damaged_products', sa.Boolean(), server_default=sa.text('false'), nullable=False))
        batch_op.add_column(sa.Column('can_add_damaged_products', sa.Boolean(), server_default=sa.text('false'), nullable=False))
        batch_op.add_column(sa.Column('can_delete_damaged_products', sa.Boolean(), server_default=sa.text('false'), nullable=False))

    # يمكن لاحقاً إزالة القيم الافتراضية إن لم نعد نحتاجها:
    for col in [
        'can_view_followups','can_add_followups','can_edit_followups','can_delete_followups',
        'can_view_damaged_products','can_add_damaged_products','can_delete_damaged_products']:
        op.execute(f"ALTER TABLE employee ALTER COLUMN {col} DROP DEFAULT;")


def downgrade():
    with op.batch_alter_table('employee') as batch_op:
        batch_op.drop_column('can_view_followups')
        batch_op.drop_column('can_add_followups')
        batch_op.drop_column('can_edit_followups')
        batch_op.drop_column('can_delete_followups')
        batch_op.drop_column('can_view_damaged_products')
        batch_op.drop_column('can_add_damaged_products')
        batch_op.drop_column('can_delete_damaged_products')
