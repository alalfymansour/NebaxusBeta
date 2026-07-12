"""add order status history table

Revision ID: 548623a7a4ca
Revises: 30e26630841c
Create Date: 2026-01-23 10:00:00.000000

"""
from alembic import op
import sqlalchemy as sa
from datetime import datetime


# revision identifiers, used by Alembic.
revision = '548623a7a4ca'
down_revision = '30e26630841c'
branch_labels = None
depends_on = None


def upgrade():
    # ===============================
    # جدول تاريخ حالات الطلبات العادية
    # ===============================
    op.create_table(
        'order_status_history',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('order_id', sa.Integer(), nullable=False),
        sa.Column('status', sa.String(length=50), nullable=False),
        sa.Column('timestamp', sa.DateTime(), nullable=False, default=datetime.utcnow),
        sa.Column('changed_by_employee_id', sa.Integer(), nullable=True),
        sa.Column('notes', sa.Text(), nullable=True),
        sa.ForeignKeyConstraint(['order_id'], ['order.id'], ondelete='CASCADE'),
        sa.ForeignKeyConstraint(['changed_by_employee_id'], ['employee.id'], ondelete='SET NULL'),
        sa.PrimaryKeyConstraint('id')
    )
    
    # إنشاء الفهارس
    op.create_index('ix_order_status_history_order_id', 'order_status_history', ['order_id'])
    op.create_index('ix_order_status_history_status', 'order_status_history', ['status'])
    op.create_index('ix_order_status_history_timestamp', 'order_status_history', ['timestamp'])
    op.create_index('ix_order_status_history_order_timestamp', 'order_status_history', ['order_id', 'timestamp'])
    op.create_index('ix_order_status_history_status_timestamp', 'order_status_history', ['status', 'timestamp'])
    
    # ملء البيانات الموجودة للطلبات العادية
    op.execute("""
        INSERT INTO order_status_history (order_id, status, timestamp, notes)
        SELECT 
            id, 
            status, 
            COALESCE(status_updated_at, date, NOW()), 
            'Initial migration record'
        FROM "order"
    """)
    
    # ===============================
    # جدول تاريخ حالات طلبات الاستبدال
    # ===============================
    op.create_table(
        'replacement_order_status_history',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('replacement_order_id', sa.Integer(), nullable=False),
        sa.Column('status', sa.String(length=50), nullable=False),
        sa.Column('timestamp', sa.DateTime(), nullable=False, default=datetime.utcnow),
        sa.Column('changed_by_employee_id', sa.Integer(), nullable=True),
        sa.Column('notes', sa.Text(), nullable=True),
        sa.ForeignKeyConstraint(['replacement_order_id'], ['replacement_order.id'], ondelete='CASCADE'),
        sa.ForeignKeyConstraint(['changed_by_employee_id'], ['employee.id'], ondelete='SET NULL'),
        sa.PrimaryKeyConstraint('id')
    )
    
    # إنشاء الفهارس لجدول طلبات الاستبدال
    op.create_index('ix_replacement_order_status_history_order_id', 'replacement_order_status_history', ['replacement_order_id'])
    op.create_index('ix_replacement_order_status_history_status', 'replacement_order_status_history', ['status'])
    op.create_index('ix_replacement_order_status_history_timestamp', 'replacement_order_status_history', ['timestamp'])
    op.create_index('ix_replacement_order_status_history_order_timestamp', 'replacement_order_status_history', ['replacement_order_id', 'timestamp'])
    op.create_index('ix_replacement_order_status_history_status_timestamp', 'replacement_order_status_history', ['status', 'timestamp'])
    
    # ملء البيانات الموجودة لطلبات الاستبدال
    op.execute("""
        INSERT INTO replacement_order_status_history (replacement_order_id, status, timestamp, notes)
        SELECT 
            id, 
            status, 
            COALESCE(status_updated_at, date, NOW()), 
            'Initial migration record'
        FROM replacement_order
    """)


def downgrade():
    # حذف فهارس جدول طلبات الاستبدال
    op.drop_index('ix_replacement_order_status_history_status_timestamp', table_name='replacement_order_status_history')
    op.drop_index('ix_replacement_order_status_history_order_timestamp', table_name='replacement_order_status_history')
    op.drop_index('ix_replacement_order_status_history_timestamp', table_name='replacement_order_status_history')
    op.drop_index('ix_replacement_order_status_history_status', table_name='replacement_order_status_history')
    op.drop_index('ix_replacement_order_status_history_order_id', table_name='replacement_order_status_history')
    
    # حذف جدول طلبات الاستبدال
    op.drop_table('replacement_order_status_history')
    
    # حذف فهارس جدول الطلبات العادية
    op.drop_index('ix_order_status_history_status_timestamp', table_name='order_status_history')
    op.drop_index('ix_order_status_history_order_timestamp', table_name='order_status_history')
    op.drop_index('ix_order_status_history_timestamp', table_name='order_status_history')
    op.drop_index('ix_order_status_history_status', table_name='order_status_history')
    op.drop_index('ix_order_status_history_order_id', table_name='order_status_history')
    
    # حذف جدول الطلبات العادية
    op.drop_table('order_status_history')
