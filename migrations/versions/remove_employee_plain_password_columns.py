"""Remove plaintext password storage columns from employee

Revision ID: rm_emp_pwd_cols_20260404
Revises: add_capital_debt_net_2026
Create Date: 2026-04-04 12:30:00.000000

"""
from alembic import op
import sqlalchemy as sa


revision = 'rm_emp_pwd_cols_20260404'
down_revision = 'add_capital_debt_net_2026'
branch_labels = None
depends_on = None


def upgrade():
    bind = op.get_bind()
    inspector = sa.inspect(bind)
    existing_columns = {col['name'] for col in inspector.get_columns('employee')}

    with op.batch_alter_table('employee', schema=None) as batch_op:
        if 'can_view_employee_passwords' in existing_columns:
            batch_op.drop_column('can_view_employee_passwords')
        if 'password_plain' in existing_columns:
            batch_op.drop_column('password_plain')


def downgrade():
    bind = op.get_bind()
    inspector = sa.inspect(bind)
    existing_columns = {col['name'] for col in inspector.get_columns('employee')}

    with op.batch_alter_table('employee', schema=None) as batch_op:
        if 'can_view_employee_passwords' not in existing_columns:
            batch_op.add_column(sa.Column('can_view_employee_passwords', sa.Boolean(), nullable=True, server_default=sa.false()))
        if 'password_plain' not in existing_columns:
            batch_op.add_column(sa.Column('password_plain', sa.String(length=255), nullable=True))
