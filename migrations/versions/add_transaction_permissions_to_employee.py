"""add transaction permissions to employee

Revision ID: def789ghi012
Revises: abc123def456
Create Date: 2025-12-12 00:01:00.000000

"""
from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = 'def789ghi012'
down_revision = 'abc123def456'
branch_labels = None
depends_on = None


def upgrade():
    # إضافة صلاحيات المعاملات للموظفين
    op.add_column('employee', sa.Column('can_add_transactions', sa.Boolean(), nullable=True))
    op.add_column('employee', sa.Column('can_edit_transactions', sa.Boolean(), nullable=True))
    op.add_column('employee', sa.Column('can_delete_transactions', sa.Boolean(), nullable=True))
    
    # تعيين القيمة الافتراضية False للصلاحيات الجديدة
    op.execute("UPDATE employee SET can_add_transactions = FALSE WHERE can_add_transactions IS NULL")
    op.execute("UPDATE employee SET can_edit_transactions = FALSE WHERE can_edit_transactions IS NULL")
    op.execute("UPDATE employee SET can_delete_transactions = FALSE WHERE can_delete_transactions IS NULL")
    
    # جعل الأعمدة غير قابلة للـ NULL
    op.alter_column('employee', 'can_add_transactions', nullable=False)
    op.alter_column('employee', 'can_edit_transactions', nullable=False)
    op.alter_column('employee', 'can_delete_transactions', nullable=False)


def downgrade():
    # حذف الصلاحيات
    op.drop_column('employee', 'can_delete_transactions')
    op.drop_column('employee', 'can_edit_transactions')
    op.drop_column('employee', 'can_add_transactions')
