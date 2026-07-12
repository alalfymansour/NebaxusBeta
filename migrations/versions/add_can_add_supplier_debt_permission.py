"""add can_add_supplier_debt permission

Revision ID: add_supplier_debt_001
Revises: merge_supplier_debt
Create Date: 2025-12-06 14:15:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'add_supplier_debt_001'
down_revision = 'merge_supplier_debt'
branch_labels = None
depends_on = None


def upgrade():
    # إنشاء جدول supplier_debts
    op.create_table('supplier_debts',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('supplier_id', sa.Integer(), nullable=False),
        sa.Column('amount', sa.Float(), nullable=False),
        sa.Column('paid_amount', sa.Float(), nullable=True),
        sa.Column('date', sa.DateTime(), nullable=True),
        sa.Column('notes', sa.Text(), nullable=True),
        sa.Column('created_by', sa.Integer(), nullable=True),
        sa.ForeignKeyConstraint(['created_by'], ['employee.id'], ),
        sa.ForeignKeyConstraint(['supplier_id'], ['supplier.id'], ),
        sa.PrimaryKeyConstraint('id')
    )
    op.create_index(op.f('ix_supplier_debts_date'), 'supplier_debts', ['date'], unique=False)
    op.create_index(op.f('ix_supplier_debts_supplier_id'), 'supplier_debts', ['supplier_id'], unique=False)
    
    # إضافة عمود can_add_supplier_debt لجدول employee
    op.add_column('employee', sa.Column('can_add_supplier_debt', sa.Boolean(), nullable=True))
    
    # تعيين القيمة الافتراضية False لجميع السجلات الموجودة
    op.execute('UPDATE employee SET can_add_supplier_debt = false WHERE can_add_supplier_debt IS NULL')
    
    # جعل العمود غير قابل للقيمة NULL
    op.alter_column('employee', 'can_add_supplier_debt', nullable=False)


def downgrade():
    # حذف عمود can_add_supplier_debt من جدول employee
    op.drop_column('employee', 'can_add_supplier_debt')
    
    # حذف جدول supplier_debts
    op.drop_index(op.f('ix_supplier_debts_supplier_id'), table_name='supplier_debts')
    op.drop_index(op.f('ix_supplier_debts_date'), table_name='supplier_debts')
    op.drop_table('supplier_debts')
