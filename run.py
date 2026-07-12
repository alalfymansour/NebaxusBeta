from app import create_app, db
from flask_migrate import Migrate
from app.models import Customer, Product, ProductVariant, Order, OrderItem, Issue, CustomerLog, FollowUp
import os

app = create_app()
migrate = Migrate(app, db)

@app.shell_context_processor
def make_shell_context():
    return {
        'db': db,
        'Customer': Customer,
        'Product': Product,
        'ProductVariant': ProductVariant,
        'Order': Order,
        'OrderItem': OrderItem,
        'Issue': Issue,
        'CustomerLog': CustomerLog,
        'FollowUp': FollowUp
    }

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)