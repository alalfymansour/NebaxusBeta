"""add new permissions: stats cards, delete_replacements, delete_salary_transactions, view_employee_passwords, password_plain

Revision ID: add_new_permissions_2026
Revises: mno789pqr012
Create Date: 2026-03-08
"""
from alembic import op
import sqlalchemy as sa

revision = 'add_new_permissions_2026'
down_revision = 'mno789pqr012'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('employee', schema=None) as batch_op:
        # صلاحيات الاستبدال المفقودة
        batch_op.add_column(sa.Column('can_delete_replacements', sa.Boolean(), nullable=True, server_default=sa.false()))
        # صلاحيات الراتب
        batch_op.add_column(sa.Column('can_delete_salary_transactions', sa.Boolean(), nullable=True, server_default=sa.false()))
        # صلاحية عرض الباسورد + عمود الباسورد النصي
        batch_op.add_column(sa.Column('can_view_employee_passwords', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('password_plain', sa.String(255), nullable=True))
        # صلاحيات كروت الإحصائيات
        batch_op.add_column(sa.Column('can_view_stats_stock', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_fixed_assets', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_total_debt', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_capital_growth', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_daily', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_pending_orders', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_net_profit', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_losses', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_delivery_rate', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_sales', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_amount_paid', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_monthly_delivered', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_monthly_invoices', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_fixed_assets_expenses', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_employee_salaries', sa.Boolean(), nullable=True, server_default=sa.false()))
        batch_op.add_column(sa.Column('can_view_stats_operational_expenses', sa.Boolean(), nullable=True, server_default=sa.false()))


def downgrade():
    with op.batch_alter_table('employee', schema=None) as batch_op:
        batch_op.drop_column('can_delete_replacements')
        batch_op.drop_column('can_delete_salary_transactions')
        batch_op.drop_column('can_view_employee_passwords')
        batch_op.drop_column('password_plain')
        batch_op.drop_column('can_view_stats_stock')
        batch_op.drop_column('can_view_stats_fixed_assets')
        batch_op.drop_column('can_view_stats_total_debt')
        batch_op.drop_column('can_view_stats_capital_growth')
        batch_op.drop_column('can_view_stats_daily')
        batch_op.drop_column('can_view_stats_pending_orders')
        batch_op.drop_column('can_view_stats_net_profit')
        batch_op.drop_column('can_view_stats_losses')
        batch_op.drop_column('can_view_stats_delivery_rate')
        batch_op.drop_column('can_view_stats_sales')
        batch_op.drop_column('can_view_stats_amount_paid')
        batch_op.drop_column('can_view_stats_monthly_delivered')
        batch_op.drop_column('can_view_stats_monthly_invoices')
        batch_op.drop_column('can_view_stats_fixed_assets_expenses')
        batch_op.drop_column('can_view_stats_employee_salaries')
        batch_op.drop_column('can_view_stats_operational_expenses')
