"""add_attendance_system

Revision ID: add_attendance_system
Revises: 25cb96bbc5ae
Create Date: 2026-06-29 12:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = 'add_attendance_system'
down_revision = '25cb96bbc5ae'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('employee', sa.Column('requires_attendance', sa.Boolean(), server_default='false', nullable=False))
    op.add_column('employee', sa.Column('can_manage_attendance', sa.Boolean(), server_default='false', nullable=False))
    op.add_column('employee', sa.Column('can_view_stats_employee_debt', sa.Boolean(), server_default='false', nullable=False))

    op.create_table(
        'attendance_record',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('employee_id', sa.Integer(), sa.ForeignKey('employee.id'), nullable=False),
        sa.Column('date', sa.Date(), nullable=False),
        sa.Column('status', sa.String(20), nullable=False),
        sa.Column('created_by', sa.Integer(), sa.ForeignKey('employee.id'), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('employee_id', 'date', name='uq_attendance_employee_date'),
    )


def downgrade():
    op.drop_table('attendance_record')
    op.drop_column('employee', 'can_view_stats_employee_debt')
    op.drop_column('employee', 'can_manage_attendance')
    op.drop_column('employee', 'requires_attendance')
