"""add cod fee support

Revision ID: cod_fee_2026_01_24
Revises: xyz123capital456
Create Date: 2026-01-24 10:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'cod_fee_2026_01_24'
down_revision = 'xyz123capital456'
branch_labels = None
depends_on = None


def upgrade():
    # إنشاء جدول app_settings
    op.create_table('app_settings',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('key', sa.String(length=100), nullable=False),
        sa.Column('value', sa.Text(), nullable=True),
        sa.Column('description', sa.String(length=255), nullable=True),
        sa.Column('updated_at', sa.DateTime(), nullable=True),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('key')
    )
    
    # إدراج القيمة الافتراضية لرسوم التحصيل
    op.execute("""
        INSERT INTO app_settings (key, value, description, updated_at)
        VALUES ('cod_fee', '0', 'رسوم التحصيل (COD Fee)', NOW())
    """)
    
    # إضافة حقل cod_fee_applied لجدول order
    op.add_column('order', sa.Column('cod_fee_applied', sa.Float(), nullable=False, server_default='0'))
    
    # إضافة حقل cod_fee_applied لجدول replacement_order
    op.add_column('replacement_order', sa.Column('cod_fee_applied', sa.Float(), nullable=False, server_default='0'))


def downgrade():
    # حذف الأعمدة
    op.drop_column('replacement_order', 'cod_fee_applied')
    op.drop_column('order', 'cod_fee_applied')
    
    # حذف الجدول
    op.drop_table('app_settings')
