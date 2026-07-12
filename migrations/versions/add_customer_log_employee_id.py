"""add employee_id to customer_log for actor attribution

Revision ID: add_customer_log_employee_id
Revises: add_order_edit_log
Create Date: 2026-03-12

"""
from alembic import op
import sqlalchemy as sa


revision = 'add_customer_log_employee_id'
down_revision = 'add_order_edit_log'
branch_labels = None
depends_on = None


def upgrade():
    conn = op.get_bind()

    result = conn.execute(sa.text("""
        SELECT column_name
        FROM information_schema.columns
        WHERE table_name = 'customer_log'
          AND column_name = 'employee_id'
    """))
    if not result.fetchone():
        op.add_column('customer_log', sa.Column('employee_id', sa.Integer(), nullable=True))

    result = conn.execute(sa.text("""
        SELECT constraint_name
        FROM information_schema.table_constraints
        WHERE table_name = 'customer_log'
          AND constraint_name = 'fk_customer_log_employee_id'
    """))
    if not result.fetchone():
        op.create_foreign_key(
            'fk_customer_log_employee_id',
            'customer_log', 'employee',
            ['employee_id'], ['id'],
            ondelete='SET NULL'
        )

    conn.execute(sa.text("""
        CREATE INDEX IF NOT EXISTS ix_customer_log_employee_id
        ON customer_log (employee_id)
    """))


def downgrade():
    op.drop_index('ix_customer_log_employee_id', table_name='customer_log')
    op.drop_constraint('fk_customer_log_employee_id', 'customer_log', type_='foreignkey')
    op.drop_column('customer_log', 'employee_id')